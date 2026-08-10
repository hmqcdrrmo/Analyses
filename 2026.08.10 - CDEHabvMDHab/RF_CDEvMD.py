r"""
Quality-control and compare the CDEHab (19-21 July 2025) and MDHab
(7-9 August 2026) hourly rainfall events.

This version uses a three-tier quality-control model:

  REJECT
      Objectively invalid observations and confirmed unusable sensors are
      omitted from calculations.

  REVIEW
      Statistically unusual but meteorologically plausible rainfall is retained
      and documented. Hourly, daily, systematic-sensor, temporal, and spatial
      tests contribute to a transparent review score.

  RETAIN
      Ordinary observations and manually approved review values remain in the
      calculations.

The supplied rain-gauge shapefile is used to compare each station with its
nearest reporting gauges. The script never silently deletes a statistical
extreme merely because it differs from the city-wide network.

Install dependencies, if necessary:

    py -m pip install pandas numpy matplotlib openpyxl geopandas shapely pyogrio

Run with the configured Windows paths:

    py rainfall_event_comparison.py

Optional overrides:

    py rainfall_event_comparison.py ^
      --cdehab "C:\path\to\CDEHab.xlsx" ^
      --mdhab "C:\path\to\MDHab.xlsx" ^
      --shapefile "C:\path\to\Rain Gauges.shp" ^
      --output "C:\path\to\Outputs"

The output directory contains raw-preserving QC tables, sensitivity statistics,
plots, a plain-language summary, and an editable manual_qc_overrides.csv file.
"""

from __future__ import annotations

import argparse
import math
import re
import sys
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

try:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
    import pandas as pd
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "A required package is missing. Run: "
        "py -m pip install pandas numpy matplotlib openpyxl geopandas shapely pyogrio"
    ) from exc


# =============================================================================
# USER CONFIGURATION
# =============================================================================

CDEHAB_FILE = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.10 - CDEHabvMDHab"
    r"\Processed Data\CDEHab.xlsx"
)

MDHAB_FILE = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.10 - CDEHabvMDHab"
    r"\Processed Data\MDHab.xlsx"
)

RAIN_GAUGE_SHAPEFILE = Path(
    r"C:\Users\QCUSER\Documents\Given\RainGauges_2026"
    r"\Rain Gauges (July 2026).shp"
)

OUTPUT_DIRECTORY = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.10 - CDEHabvMDHab\Outputs"
)


@dataclass(frozen=True)
class QCSettings:
    """Auditable quality-control settings."""

    # Hard rejection: only clearly invalid data.
    hard_max_hourly_mm: float = 300.0

    # Network and hourly review tests.
    minimum_hourly_review_mm: float = 7.5
    extreme_hourly_review_mm: float = 75.0
    modified_z_review: float = 3.0
    modified_z_strong: float = 3.5
    hourly_iqr_soft_multiplier: float = 1.5
    hourly_iqr_strong_multiplier: float = 3.0
    minimum_reporting_sensors: int = 8

    # Temporal review.
    temporal_window_hours: int = 2
    temporal_ratio: float = 3.0
    temporal_difference_mm: float = 15.0

    # Spatial review.
    spatial_radius_km: float = 3.0
    spatial_min_neighbors: int = 3
    spatial_max_neighbors: int = 5
    spatial_ratio: float = 2.5
    spatial_difference_mm: float = 10.0
    spatial_match_minimum_score: float = 0.80

    # Sensor status and systematic behavior.
    network_wet_threshold_mm: float = 2.5
    minimum_network_wet_hours: int = 3
    inactive_zero_fraction: float = 0.90
    constant_sensor_minimum_coverage: float = 0.90
    systematic_minimum_wet_hours: int = 6
    systematic_median_ratio: float = 1.50
    systematic_high_fraction: float = 0.25
    systematic_high_minimum_difference_mm: float = 5.0
    systematic_low_correlation: float = 0.30
    systematic_minimum_mean_difference_mm: float = 10.0
    lag_search_hours: int = 2
    lag_minimum_correlation: float = 0.50
    lag_minimum_improvement: float = 0.15

    # Daily completeness and review.
    minimum_daily_completeness: float = 0.90
    disqualify_missing_wet_hours: bool = True
    daily_iqr_multiplier: float = 1.5
    daily_modified_z_review: float = 3.5
    daily_spatial_ratio: float = 1.75
    daily_spatial_difference_mm: float = 30.0
    trimmed_mean_fraction: float = 0.10

    # Scoring.
    review_score_threshold: int = 2
    severe_review_score_threshold: int = 4


DEFAULT_SETTINGS = QCSettings()

EVENT_COLORS = {"CDEHab": "#2878B5", "MDHab": "#E07A1F"}

HOURLY_INTENSITY_LINES = (
    (2.5, "Moderate", "#F39C12"),
    (7.5, "Heavy", "#D62728"),
    (15.0, "Intense", "#8B0000"),
    (30.0, "Torrential", "#7B2CBF"),
)

# These aliases are applied before fuzzy matching. Add or correct entries here
# if station_spatial_matches.csv shows a wrong or unmatched gauge.
MANUAL_STATION_ALIASES: dict[str, str] = {
    "Brgy. Baesa": "Brgy Baesa Hall",
    "Brgy. Fairview": "Brgy Fairview (REC)",
    "Brgy. NS Amoranto": "Brgy N.S Amoranto Hall",
    "Brgy. Valencia": "Brgy Valencia Hall",
    "Commonwealth Highschool": "Commonwealth High School",
    "Dona Juana Elementary School": "Dona Juana Elementary School",
    "Doña Imelda": "Dona Imelda",
    "Emilio Jacinto": "Emilio Jacinto Sr HS",
    "Kaingin 1 Satellite Office": "Pansol Kaingin 1 Brgy Satellite Office",
    "Laging Handa Barangay Hall": "Laging Handa Hall",
    "Payatas Elementary School": "Payatas ES",
    "Pinyahan Multipurpose Hall": "Pinyahan Multipurose Hall",
    "Placido Delmundo Elementary School": "Placido Del Mundo Elementary School",
    "Ramon Magsaysay": "Ramon Magsaysay HS",
    "Ramon Magsaysay Brgy Hall": "Ramon Magsaysay Brgy Hall",
    "Ramon Magsaysay Elementary School": "Ramon Magsaysay Elementary School",
}

MANUAL_OVERRIDE_COLUMNS = ["event", "timestamp", "station", "action", "reason"]

INVALID_COLUMNS = [
    "event",
    "timestamp",
    "station",
    "original_value",
    "reason",
    "action",
]

HOURLY_REVIEW_COLUMNS = [
    "event",
    "timestamp",
    "station",
    "value_mm",
    "review_level",
    "review_score",
    "reasons",
    "network_median_mm",
    "modified_z",
    "iqr_soft_upper_mm",
    "iqr_strong_upper_mm",
    "temporal_median_mm",
    "neighbor_median_mm",
    "neighbor_count",
    "neighbor_ratio",
    "spatial_match_available",
    "systematic_sensor_candidate",
]

DAILY_REVIEW_COLUMNS = [
    "event",
    "day_number",
    "date",
    "station",
    "daily_accumulation_mm",
    "review_level",
    "review_score",
    "reasons",
    "network_daily_median_mm",
    "modified_z",
    "iqr_upper_mm",
    "neighbor_daily_median_mm",
    "neighbor_count",
]

DATA_ISSUE_COLUMNS = ["event", "issue", "details"]


# =============================================================================
# GENERAL HELPERS
# =============================================================================


def parse_datetime_cell(value: Any) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.Timestamp(value)
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
        except (ValueError, TypeError, OverflowError):
            return pd.NaT

    text = re.sub(r"\s+", " ", str(value).strip())
    for date_format in (
        "%d %m %Y %H:%M",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ):
        try:
            return pd.Timestamp(datetime.strptime(text, date_format))
        except ValueError:
            pass
    return pd.to_datetime(text, dayfirst=True, errors="coerce")


def parse_rainfall_cell(value: Any) -> float:
    if pd.isna(value):
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        return float(value)
    text = str(value).strip()
    if text.casefold() in {"", "-", "--", "na", "n/a", "nan", "none"}:
        return np.nan
    text = text.replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return np.nan


def is_meaningful_unparsed_value(value: Any) -> bool:
    if pd.isna(value):
        return False
    return str(value).strip().casefold() not in {
        "",
        "-",
        "--",
        "na",
        "n/a",
        "nan",
        "none",
    }


def find_datetime_column(columns: list[Any]) -> Any:
    normalized = {
        str(column).strip().casefold().replace("_", " "): column for column in columns
    }
    for candidate in (
        "date & time",
        "date and time",
        "date time",
        "datetime",
        "date pht",
        "date",
    ):
        if candidate in normalized:
            return normalized[candidate]
    return columns[0]


def normalize_name(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(character for character in text if not unicodedata.combining(character))
    text = text.casefold().replace("&", " and ")
    text = re.sub(r"\bbarangay\b", " brgy ", text)
    text = re.sub(r"\bsenior high school\b", " shs ", text)
    text = re.sub(r"\bsr\.?\s*high school\b", " shs ", text)
    text = re.sub(r"\bhighschool\b|\bhigh school\b", " hs ", text)
    text = re.sub(r"\belementary school\b", " es ", text)
    text = re.sub(r"\bmultipurose\b", " multipurpose ", text)
    text = re.sub(r"\bquezon city university\b", " qcu ", text)
    text = re.sub(r"\bbarangay hall\b", " brgy hall ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def station_similarity(left: str, right: str) -> float:
    a = normalize_name(left)
    b = normalize_name(right)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    sequence = SequenceMatcher(None, a, b).ratio()
    tokens_a, tokens_b = set(a.split()), set(b.split())
    union = tokens_a | tokens_b
    jaccard = len(tokens_a & tokens_b) / len(union) if union else 0.0
    containment = (
        len(tokens_a & tokens_b) / min(len(tokens_a), len(tokens_b))
        if tokens_a and tokens_b
        else 0.0
    )
    substring = 0.92 if (a in b or b in a) and min(len(a), len(b)) >= 6 else 0.0
    return max(sequence, jaccard, 0.95 * containment, substring)


def robust_distribution_metrics(values: Iterable[float]) -> dict[str, float]:
    series = pd.Series(values, dtype=float).dropna()
    if series.empty:
        return {
            "median": np.nan,
            "mad": np.nan,
            "robust_sigma": np.nan,
            "q1": np.nan,
            "q3": np.nan,
            "iqr": np.nan,
        }
    median = float(series.median())
    mad = float((series - median).abs().median())
    q1 = float(series.quantile(0.25))
    q3 = float(series.quantile(0.75))
    return {
        "median": median,
        "mad": mad,
        "robust_sigma": 1.4826 * mad,
        "q1": q1,
        "q3": q3,
        "iqr": q3 - q1,
    }


def modified_z(value: float, median: float, mad: float) -> float:
    if pd.isna(value) or pd.isna(median) or pd.isna(mad):
        return np.nan
    if mad == 0:
        if value > median:
            return np.inf
        if value < median:
            return -np.inf
        return 0.0
    return 0.6745 * (value - median) / mad


def trimmed_mean(values: Iterable[float], fraction: float) -> float:
    array = np.sort(pd.Series(values, dtype=float).dropna().to_numpy())
    if len(array) == 0:
        return np.nan
    trim = int(math.floor(len(array) * fraction))
    if 2 * trim >= len(array):
        return float(np.mean(array))
    return float(np.mean(array[trim : len(array) - trim]))


def safe_ratio(numerator: float, denominator: float, offset: float = 0.5) -> float:
    if pd.isna(numerator) or pd.isna(denominator):
        return np.nan
    return float((numerator + offset) / (denominator + offset))


def safe_correlation(left: pd.Series, right: pd.Series, minimum_count: int = 4) -> float:
    pair = pd.concat([left, right], axis=1).dropna()
    if len(pair) < minimum_count:
        return np.nan
    if pair.iloc[:, 0].nunique() < 2 or pair.iloc[:, 1].nunique() < 2:
        return np.nan
    return float(pair.iloc[:, 0].corr(pair.iloc[:, 1]))


def prepare_csv_frame(frame: pd.DataFrame) -> pd.DataFrame:
    output = frame.copy()
    for column in output.columns:
        if pd.api.types.is_datetime64_any_dtype(output[column]):
            if "timestamp" in column:
                output[column] = output[column].dt.strftime("%Y-%m-%d %H:%M")
            else:
                output[column] = output[column].dt.strftime("%Y-%m-%d")
    return output


def save_csv(frame: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    prepare_csv_frame(frame).to_csv(
        path,
        index=False,
        encoding="utf-8-sig",
        float_format="%.3f",
    )


# =============================================================================
# EXCEL INPUT AND HARD REJECTION
# =============================================================================


def load_event(
    event: str,
    file_path: Path,
    settings: QCSettings,
) -> tuple[pd.DataFrame, list[dict[str, Any]], list[dict[str, str]]]:
    if not file_path.exists():
        raise FileNotFoundError(f"Input workbook not found for {event}:\n{file_path}")

    try:
        raw = pd.read_excel(file_path, sheet_name=0, dtype=object, engine="openpyxl")
    except ImportError as exc:
        raise RuntimeError(
            "Reading .xlsx files requires openpyxl. Run: py -m pip install openpyxl"
        ) from exc

    raw = raw.dropna(axis=1, how="all")
    if raw.empty or raw.shape[1] < 2:
        raise ValueError(f"{file_path.name} does not contain a usable rainfall table.")

    datetime_column = find_datetime_column(list(raw.columns))
    timestamps = raw[datetime_column].map(parse_datetime_cell)
    invalid_timestamp_count = int(timestamps.isna().sum())
    issues: list[dict[str, str]] = []
    if invalid_timestamp_count:
        issues.append(
            {
                "event": event,
                "issue": "non_data_or_invalid_timestamp_rows_removed",
                "details": str(invalid_timestamp_count),
            }
        )

    valid_rows = timestamps.notna()
    raw = raw.loc[valid_rows].copy()
    timestamps = timestamps.loc[valid_rows]
    if raw.empty:
        raise ValueError(f"No valid timestamps were found in {file_path.name}.")

    station_columns = [column for column in raw.columns if column != datetime_column]
    original = raw[station_columns].copy()
    numeric = pd.DataFrame(index=raw.index)
    rejected: list[dict[str, Any]] = []

    for original_station in station_columns:
        station = str(original_station).strip()
        parsed = original[original_station].map(parse_rainfall_cell)
        numeric[station] = parsed
        malformed = parsed.isna() & original[original_station].map(
            is_meaningful_unparsed_value
        )
        for row_index in original.index[malformed]:
            rejected.append(
                {
                    "event": event,
                    "timestamp": timestamps.loc[row_index],
                    "station": station,
                    "original_value": original.at[row_index, original_station],
                    "reason": "non_numeric_value",
                    "action": "rejected",
                }
            )

    for row_index, timestamp in timestamps.items():
        for station in numeric.columns:
            value = numeric.at[row_index, station]
            if pd.isna(value):
                continue
            reason = None
            if value < 0:
                reason = "negative_rainfall"
            elif value > settings.hard_max_hourly_mm:
                reason = f"above_hard_limit_{settings.hard_max_hourly_mm:g}_mm"
            if reason:
                rejected.append(
                    {
                        "event": event,
                        "timestamp": timestamp,
                        "station": station,
                        "original_value": value,
                        "reason": reason,
                        "action": "rejected",
                    }
                )
                numeric.at[row_index, station] = np.nan

    numeric.index = pd.DatetimeIndex(timestamps.values, name="Date & Time")
    duplicate_count = int(numeric.index.duplicated(keep=False).sum())
    if duplicate_count:
        issues.append(
            {
                "event": event,
                "issue": "duplicate_timestamp_rows_averaged",
                "details": str(duplicate_count),
            }
        )
        numeric = numeric.groupby(level=0).mean()

    numeric = numeric.sort_index()
    start = numeric.index.min().floor("D")
    end = numeric.index.max().floor("D") + pd.Timedelta(hours=23)
    full_index = pd.date_range(start, end, freq="h", name="Date & Time")
    missing_timestamps = int(len(full_index.difference(numeric.index)))
    if missing_timestamps:
        issues.append(
            {
                "event": event,
                "issue": "missing_hourly_timestamps_inserted",
                "details": str(missing_timestamps),
            }
        )
    return numeric.reindex(full_index), rejected, issues


# =============================================================================
# SENSOR STATUS AND MANUAL OVERRIDES
# =============================================================================


def load_manual_overrides(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=MANUAL_OVERRIDE_COLUMNS)
    frame = pd.read_csv(path, dtype=str).fillna("")
    missing = [column for column in MANUAL_OVERRIDE_COLUMNS if column not in frame]
    if missing:
        raise ValueError(
            f"{path.name} is missing required columns: {', '.join(missing)}"
        )
    frame = frame[MANUAL_OVERRIDE_COLUMNS].copy()
    frame["action"] = frame["action"].str.strip().str.casefold()
    return frame.loc[frame["action"] != ""].copy()


def preliminary_reference_columns(rainfall: pd.DataFrame) -> list[str]:
    columns: list[str] = []
    for station in rainfall:
        valid = rainfall[station].dropna()
        if not valid.empty and (valid > 0).any() and valid.nunique() > 1:
            columns.append(station)
    if not columns:
        columns = [station for station in rainfall if rainfall[station].notna().any()]
    return columns


def identify_sensor_status(
    event: str,
    rainfall: pd.DataFrame,
    settings: QCSettings,
    overrides: pd.DataFrame,
) -> pd.DataFrame:
    reference_columns = preliminary_reference_columns(rainfall)
    expected = len(rainfall)
    rows: list[dict[str, Any]] = []

    for station in rainfall.columns:
        series = rainfall[station]
        valid = series.dropna()
        other_reference = [name for name in reference_columns if name != station]
        network = (
            rainfall[other_reference].median(axis=1, skipna=True)
            if other_reference
            else pd.Series(np.nan, index=rainfall.index)
        )
        network_wet = network >= settings.network_wet_threshold_mm
        network_wet_hours = int(network_wet.sum())
        comparable = network_wet & series.notna()
        zero_when_wet_fraction = (
            float((series.loc[comparable] == 0).mean()) if comparable.any() else np.nan
        )

        positive_count = int((valid > 0).sum())
        included = True
        status = "active"
        reason = ""

        if valid.empty:
            included = False
            status = "all_missing"
            reason = "No usable observations"
        elif positive_count == 0:
            if (
                network_wet_hours >= settings.minimum_network_wet_hours
                and zero_when_wet_fraction >= settings.inactive_zero_fraction
            ):
                included = False
                status = "all_zero_while_network_wet"
                reason = "Likely inactive sensor"
            else:
                status = "all_zero_unconfirmed"
                reason = "Retained because the event lacks enough wet-network evidence"
        elif (
            valid.nunique() == 1
            and len(valid) >= math.ceil(settings.constant_sensor_minimum_coverage * expected)
        ):
            included = False
            status = "constant_value"
            reason = "Likely stuck sensor"
        elif (
            network_wet_hours >= settings.minimum_network_wet_hours
            and zero_when_wet_fraction >= settings.inactive_zero_fraction
        ):
            status = "mostly_zero_when_network_wet_review"
            reason = "Retained but requires sensor-status review"

        event_overrides = overrides.loc[
            (overrides["event"].str.casefold() == event.casefold())
            & (overrides["station"].str.casefold() == station.casefold())
        ]
        for _, override in event_overrides.iterrows():
            if override["action"] == "exclude_sensor":
                included = False
                status = "manual_exclusion"
                reason = override["reason"] or "Manual sensor exclusion"
            elif override["action"] == "include_sensor":
                included = True
                status = "manual_inclusion"
                reason = override["reason"] or "Manual sensor inclusion"

        rows.append(
            {
                "event": event,
                "station": station,
                "status": status,
                "status_reason": reason,
                "included_in_analysis": included,
                "valid_observations": int(valid.count()),
                "expected_observations": expected,
                "positive_observations": positive_count,
                "network_wet_hours": network_wet_hours,
                "zero_when_network_wet_fraction": zero_when_wet_fraction,
                "minimum_mm": valid.min() if not valid.empty else np.nan,
                "maximum_mm": valid.max() if not valid.empty else np.nan,
            }
        )
    return pd.DataFrame(rows)


def apply_manual_hour_actions(
    event: str,
    rainfall: pd.DataFrame,
    overrides: pd.DataFrame,
) -> tuple[pd.DataFrame, list[dict[str, Any]], set[tuple[pd.Timestamp, str]], pd.DataFrame]:
    cleaned = rainfall.copy()
    rejected: list[dict[str, Any]] = []
    retained: set[tuple[pd.Timestamp, str]] = set()
    action_rows: list[dict[str, Any]] = []

    event_rows = overrides.loc[
        overrides["event"].str.casefold() == event.casefold()
    ]
    for _, override in event_rows.iterrows():
        action = override["action"]
        if action not in {"reject_hour", "retain_hour"}:
            continue
        timestamp = parse_datetime_cell(override["timestamp"])
        station = override["station"].strip()
        applied = False
        if timestamp in cleaned.index and station in cleaned.columns:
            value = cleaned.at[timestamp, station]
            if action == "reject_hour":
                rejected.append(
                    {
                        "event": event,
                        "timestamp": timestamp,
                        "station": station,
                        "original_value": value,
                        "reason": override["reason"] or "manual_hour_rejection",
                        "action": "rejected",
                    }
                )
                cleaned.at[timestamp, station] = np.nan
            else:
                retained.add((pd.Timestamp(timestamp), station))
            applied = True
        action_rows.append(
            {
                "event": event,
                "timestamp": timestamp,
                "station": station,
                "action": action,
                "reason": override["reason"],
                "applied": applied,
            }
        )
    return cleaned, rejected, retained, pd.DataFrame(action_rows)


# =============================================================================
# SHAPEFILE MATCHING AND NEIGHBORS
# =============================================================================


def select_gauge_name_field(gdf: Any, sensor_names: list[str]) -> str:
    candidate_fields = [
        column
        for column in gdf.columns
        if column != gdf.geometry.name
        and (
            pd.api.types.is_object_dtype(gdf[column])
            or pd.api.types.is_string_dtype(gdf[column])
        )
        and gdf[column].notna().any()
    ]
    if not candidate_fields:
        raise ValueError("No text field was found in the rain-gauge shapefile.")

    preferred_words = {"name", "station", "gauge", "location", "site"}
    field_scores: list[tuple[float, str]] = []
    sample_sensors = sensor_names[: min(len(sensor_names), 100)]
    for field in candidate_fields:
        values = gdf[field].dropna().astype(str).tolist()
        if not values:
            continue
        best_scores = [
            max(station_similarity(sensor, value) for value in values)
            for sensor in sample_sensors
        ]
        preference = 0.05 if any(word in field.casefold() for word in preferred_words) else 0
        field_scores.append((float(np.mean(best_scores)) + preference, field))
    if not field_scores:
        raise ValueError("The shapefile text fields contain no usable gauge names.")
    return max(field_scores)[1]


def load_gauge_locations(
    shapefile_path: Path,
    sensor_names: list[str],
    explicit_name_field: str | None,
) -> tuple[pd.DataFrame, str]:
    if not shapefile_path.exists():
        raise FileNotFoundError(f"Rain-gauge shapefile not found:\n{shapefile_path}")
    try:
        import geopandas as gpd
    except ImportError as exc:
        raise RuntimeError(
            "Spatial verification requires geopandas. Run: "
            "py -m pip install geopandas shapely pyogrio"
        ) from exc

    gauges = gpd.read_file(shapefile_path)
    if gauges.empty:
        raise ValueError("The rain-gauge shapefile contains no features.")
    if gauges.crs is None:
        raise ValueError(
            "The rain-gauge shapefile has no CRS. Define its CRS before running "
            "spatial verification."
        )
    if explicit_name_field:
        if explicit_name_field not in gauges.columns:
            raise ValueError(
                f"Gauge-name field '{explicit_name_field}' is not in the shapefile."
            )
        name_field = explicit_name_field
    else:
        name_field = select_gauge_name_field(gauges, sensor_names)

    projected = gauges.to_crs(epsg=32651)
    geometry = projected.geometry
    point_geometry = geometry.where(geometry.geom_type == "Point", geometry.centroid)
    valid = point_geometry.notna() & ~point_geometry.is_empty & projected[name_field].notna()
    table = pd.DataFrame(
        {
            "gauge_index": projected.index[valid].astype(str),
            "gauge_name": projected.loc[valid, name_field].astype(str).str.strip().values,
            "x_m": point_geometry.loc[valid].x.values,
            "y_m": point_geometry.loc[valid].y.values,
        }
    )
    if table.empty:
        raise ValueError("No valid named geometries were found in the shapefile.")
    return table.reset_index(drop=True), name_field


def match_stations_to_gauges(
    events_and_stations: dict[str, list[str]],
    gauge_locations: pd.DataFrame,
    settings: QCSettings,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    gauge_names = gauge_locations["gauge_name"].tolist()
    normalized_lookup: dict[str, list[int]] = {}
    for index, gauge_name in enumerate(gauge_names):
        normalized_lookup.setdefault(normalize_name(gauge_name), []).append(index)

    for event, stations in events_and_stations.items():
        for station in stations:
            alias = MANUAL_STATION_ALIASES.get(station, station)
            normalized_alias = normalize_name(alias)
            if normalized_alias in normalized_lookup:
                gauge_index = normalized_lookup[normalized_alias][0]
                score = 1.0
                method = "manual_alias_exact" if alias != station else "exact"
            else:
                scores = [station_similarity(alias, gauge_name) for gauge_name in gauge_names]
                gauge_index = int(np.argmax(scores))
                score = float(scores[gauge_index])
                method = "fuzzy"

            matched = score >= settings.spatial_match_minimum_score
            gauge = gauge_locations.iloc[gauge_index] if matched else None
            rows.append(
                {
                    "event": event,
                    "station": station,
                    "alias_used": alias,
                    "matched": matched,
                    "matched_gauge_name": gauge["gauge_name"] if matched else "",
                    "match_score": score,
                    "match_method": method if matched else "unmatched",
                    "x_m": gauge["x_m"] if matched else np.nan,
                    "y_m": gauge["y_m"] if matched else np.nan,
                }
            )
    return pd.DataFrame(rows)


def build_neighbor_table(
    spatial_matches: pd.DataFrame,
    sensor_status: pd.DataFrame,
    settings: QCSettings,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for event, event_matches in spatial_matches.groupby("event", sort=False):
        included = sensor_status.loc[
            (sensor_status["event"] == event)
            & sensor_status["included_in_analysis"],
            "station",
        ]
        matched = event_matches.loc[
            event_matches["matched"] & event_matches["station"].isin(included)
        ].copy()
        for _, station_row in matched.iterrows():
            other = matched.loc[matched["station"] != station_row["station"]].copy()
            if other.empty:
                continue
            other["distance_km"] = np.sqrt(
                (other["x_m"] - station_row["x_m"]) ** 2
                + (other["y_m"] - station_row["y_m"]) ** 2
            ) / 1000.0
            other = other.loc[other["distance_km"] > 0.001].sort_values("distance_km")
            within = other.loc[
                other["distance_km"] <= settings.spatial_radius_km
            ].head(settings.spatial_max_neighbors)
            chosen = within.copy()
            if len(chosen) < settings.spatial_min_neighbors:
                needed = settings.spatial_min_neighbors - len(chosen)
                supplement = other.loc[~other.index.isin(chosen.index)].head(needed)
                chosen = pd.concat([chosen, supplement]).head(
                    settings.spatial_max_neighbors
                )
            for rank, (_, neighbor) in enumerate(chosen.iterrows(), start=1):
                rows.append(
                    {
                        "event": event,
                        "station": station_row["station"],
                        "neighbor_station": neighbor["station"],
                        "neighbor_rank": rank,
                        "distance_km": neighbor["distance_km"],
                        "within_configured_radius": (
                            neighbor["distance_km"] <= settings.spatial_radius_km
                        ),
                    }
                )
    return pd.DataFrame(
        rows,
        columns=[
            "event",
            "station",
            "neighbor_station",
            "neighbor_rank",
            "distance_km",
            "within_configured_radius",
        ],
    )


def neighbor_map_from_table(
    neighbor_table: pd.DataFrame,
) -> dict[str, dict[str, list[str]]]:
    mapping: dict[str, dict[str, list[str]]] = {}
    for (event, station), group in neighbor_table.groupby(
        ["event", "station"], sort=False
    ):
        mapping.setdefault(event, {})[station] = group.sort_values(
            "neighbor_rank"
        )["neighbor_station"].tolist()
    return mapping


# =============================================================================
# SYSTEMATIC SENSOR DIAGNOSTICS
# =============================================================================


def calculate_sensor_diagnostics(
    event: str,
    rainfall: pd.DataFrame,
    active_stations: list[str],
    neighbors: dict[str, list[str]],
    settings: QCSettings,
) -> pd.DataFrame:
    active = rainfall[active_stations]
    rows: list[dict[str, Any]] = []

    for station in active_stations:
        series = active[station]
        other = [name for name in active_stations if name != station]
        network = active[other].median(axis=1, skipna=True)
        wet = network >= settings.network_wet_threshold_mm
        comparable = wet & series.notna() & network.notna()
        wet_count = int(comparable.sum())
        ratios = (series.loc[comparable] + 0.5) / (network.loc[comparable] + 0.5)
        high = (
            (series.loc[comparable] - network.loc[comparable] >= 5.0)
            & (ratios >= 2.0)
        )
        zero_fraction = (
            float((series.loc[comparable] == 0).mean()) if wet_count else np.nan
        )
        median_ratio = float(ratios.median()) if wet_count else np.nan
        high_fraction = float(high.mean()) if wet_count else np.nan
        mean_difference = (
            float((series.loc[comparable] - network.loc[comparable]).abs().mean())
            if wet_count
            else np.nan
        )
        same_hour_corr = safe_correlation(
            series.loc[comparable], network.loc[comparable]
        )

        lag_correlations: dict[int, float] = {}
        for lag in range(-settings.lag_search_hours, settings.lag_search_hours + 1):
            shifted_network = network.shift(lag)
            lag_mask = wet & series.notna() & shifted_network.notna()
            lag_correlations[lag] = safe_correlation(
                series.loc[lag_mask], shifted_network.loc[lag_mask]
            )
        valid_lags = {
            lag: correlation
            for lag, correlation in lag_correlations.items()
            if not pd.isna(correlation)
        }
        if valid_lags:
            best_lag = max(valid_lags, key=valid_lags.get)
            best_lag_corr = valid_lags[best_lag]
        else:
            best_lag = 0
            best_lag_corr = np.nan

        neighbor_names = [name for name in neighbors.get(station, []) if name in active]
        neighbor_median = (
            active[neighbor_names].median(axis=1, skipna=True)
            if neighbor_names
            else pd.Series(np.nan, index=active.index)
        )
        neighbor_wet = neighbor_median >= settings.network_wet_threshold_mm
        neighbor_comparable = neighbor_wet & series.notna()
        neighbor_ratios = (
            (series.loc[neighbor_comparable] + 0.5)
            / (neighbor_median.loc[neighbor_comparable] + 0.5)
            if neighbor_comparable.any()
            else pd.Series(dtype=float)
        )
        median_neighbor_ratio = (
            float(neighbor_ratios.median()) if not neighbor_ratios.empty else np.nan
        )

        reasons: list[str] = []
        systematic_high = (
            wet_count >= settings.systematic_minimum_wet_hours
            and (
                (
                    not pd.isna(median_ratio)
                    and median_ratio >= settings.systematic_median_ratio
                    and high_fraction >= settings.systematic_high_fraction
                )
                or (
                    high_fraction >= settings.systematic_high_fraction
                    and mean_difference
                    >= settings.systematic_high_minimum_difference_mm
                )
            )
        )
        low_correlation = (
            wet_count >= settings.systematic_minimum_wet_hours
            and not pd.isna(same_hour_corr)
            and same_hour_corr < settings.systematic_low_correlation
            and mean_difference >= settings.systematic_minimum_mean_difference_mm
        )
        lag_improvement = (
            best_lag_corr - same_hour_corr
            if not pd.isna(best_lag_corr) and not pd.isna(same_hour_corr)
            else np.nan
        )
        possible_time_shift = (
            best_lag != 0
            and not pd.isna(best_lag_corr)
            and best_lag_corr >= settings.lag_minimum_correlation
            and not pd.isna(lag_improvement)
            and lag_improvement >= settings.lag_minimum_improvement
        )
        if systematic_high:
            reasons.append("persistent_high_bias")
        if low_correlation:
            reasons.append("low_network_correlation")
        if possible_time_shift:
            reasons.append(f"possible_time_shift_{best_lag:+d}h")
        if (
            not pd.isna(median_neighbor_ratio)
            and median_neighbor_ratio >= settings.systematic_median_ratio
            and len(neighbor_names) >= settings.spatial_min_neighbors
        ):
            reasons.append("persistent_neighbor_disagreement")

        rows.append(
            {
                "event": event,
                "station": station,
                "wet_comparison_hours": wet_count,
                "same_hour_network_correlation": same_hour_corr,
                "median_network_ratio": median_ratio,
                "high_ratio_fraction": high_fraction,
                "zero_when_network_wet_fraction": zero_fraction,
                "mean_absolute_network_difference_mm": mean_difference,
                "best_lag_hours": best_lag,
                "best_lag_correlation": best_lag_corr,
                "lag_correlation_improvement": lag_improvement,
                "spatial_neighbor_count": len(neighbor_names),
                "median_neighbor_ratio": median_neighbor_ratio,
                "systematic_review_candidate": bool(reasons),
                "review_reasons": ";".join(reasons),
            }
        )
    return pd.DataFrame(rows)


# =============================================================================
# HOURLY REVIEW SCORING
# =============================================================================


def score_hourly_reviews(
    event: str,
    rainfall: pd.DataFrame,
    active_stations: list[str],
    neighbors: dict[str, list[str]],
    diagnostics: pd.DataFrame,
    manual_retained: set[tuple[pd.Timestamp, str]],
    settings: QCSettings,
) -> pd.DataFrame:
    active = rainfall[active_stations]
    systematic_lookup = diagnostics.set_index("station")[
        "systematic_review_candidate"
    ].to_dict()
    rows: list[dict[str, Any]] = []

    for timestamp, hour_values in active.iterrows():
        for station, value in hour_values.dropna().items():
            if (pd.Timestamp(timestamp), station) in manual_retained:
                continue
            other_values = hour_values.drop(labels=[station]).dropna()
            if len(other_values) < settings.minimum_reporting_sensors:
                continue
            metrics = robust_distribution_metrics(other_values)
            soft_upper = metrics["q3"] + settings.hourly_iqr_soft_multiplier * metrics["iqr"]
            strong_upper = (
                metrics["q3"] + settings.hourly_iqr_strong_multiplier * metrics["iqr"]
            )
            z_score = modified_z(value, metrics["median"], metrics["mad"])

            score = 0
            reasons: list[str] = []
            eligible = value >= settings.minimum_hourly_review_mm
            if eligible and value > soft_upper:
                score += 1
                reasons.append("above_1.5_iqr")
            if eligible and value > strong_upper:
                score += 1
                reasons.append("above_3_iqr")
            if eligible and z_score > settings.modified_z_review:
                score += 1
                reasons.append("modified_z_above_3")
            if eligible and z_score > settings.modified_z_strong:
                reasons.append("modified_z_above_3.5")
            if value >= settings.extreme_hourly_review_mm:
                score += 1
                reasons.append(f"extreme_at_least_{settings.extreme_hourly_review_mm:g}_mm")

            window = active.loc[
                timestamp - pd.Timedelta(hours=settings.temporal_window_hours) :
                timestamp + pd.Timedelta(hours=settings.temporal_window_hours),
                station,
            ]
            window = window.loc[window.index != timestamp].dropna()
            temporal_median = float(window.median()) if not window.empty else np.nan
            temporal_isolated = (
                not pd.isna(temporal_median)
                and value - temporal_median >= settings.temporal_difference_mm
                and safe_ratio(value, temporal_median)
                >= settings.temporal_ratio
            )
            if temporal_isolated:
                score += 1
                reasons.append("temporally_isolated")

            neighbor_names = [
                name for name in neighbors.get(station, []) if name in active.columns
            ]
            neighbor_values = hour_values[neighbor_names].dropna()
            neighbor_count = len(neighbor_values)
            neighbor_median = (
                float(neighbor_values.median()) if neighbor_count else np.nan
            )
            neighbor_ratio = safe_ratio(value, neighbor_median)
            spatial_available = neighbor_count >= settings.spatial_min_neighbors
            spatial_disagreement = (
                spatial_available
                and value - neighbor_median >= settings.spatial_difference_mm
                and neighbor_ratio >= settings.spatial_ratio
            )
            if spatial_disagreement:
                score += 2
                reasons.append("nearest_neighbor_disagreement")

            systematic = bool(systematic_lookup.get(station, False))
            if systematic and eligible and score > 0:
                score += 1
                reasons.append("sensor_has_systematic_review_flag")

            if (
                score >= settings.review_score_threshold
                or value >= max(100.0, settings.extreme_hourly_review_mm)
            ):
                level = (
                    "severe_review"
                    if score >= settings.severe_review_score_threshold
                    else "review"
                )
                rows.append(
                    {
                        "event": event,
                        "timestamp": timestamp,
                        "station": station,
                        "value_mm": value,
                        "review_level": level,
                        "review_score": score,
                        "reasons": ";".join(reasons),
                        "network_median_mm": metrics["median"],
                        "modified_z": z_score,
                        "iqr_soft_upper_mm": soft_upper,
                        "iqr_strong_upper_mm": strong_upper,
                        "temporal_median_mm": temporal_median,
                        "neighbor_median_mm": neighbor_median,
                        "neighbor_count": neighbor_count,
                        "neighbor_ratio": neighbor_ratio,
                        "spatial_match_available": spatial_available,
                        "systematic_sensor_candidate": systematic,
                    }
                )
    return pd.DataFrame(rows, columns=HOURLY_REVIEW_COLUMNS)


# =============================================================================
# DAILY AGGREGATION, DAILY REVIEW, AND SENSITIVITY
# =============================================================================


def calculate_daily_statistics(
    event: str,
    rainfall: pd.DataFrame,
    active_stations: list[str],
    neighbors: dict[str, list[str]],
    diagnostics: pd.DataFrame,
    settings: QCSettings,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    active = rainfall[active_stations]
    start_date = active.index.min().floor("D")
    dates = pd.date_range(start_date, active.index.max().floor("D"), freq="D")
    minimum_valid = math.ceil(24 * settings.minimum_daily_completeness)
    sensor_rows: list[dict[str, Any]] = []

    for date in dates:
        day = active.loc[date : date + pd.Timedelta(hours=23)]
        day_network = day.median(axis=1, skipna=True)
        network_wet = day_network >= settings.network_wet_threshold_mm
        for station in active_stations:
            valid_hours = int(day[station].count())
            missing_wet_hours = int((day[station].isna() & network_wet).sum())
            qualifies = valid_hours >= minimum_valid and (
                not settings.disqualify_missing_wet_hours or missing_wet_hours == 0
            )
            sensor_rows.append(
                {
                    "event": event,
                    "day_number": int((date - dates[0]).days + 1),
                    "date": date,
                    "station": station,
                    "daily_accumulation_mm": (
                        float(day[station].sum()) if qualifies else np.nan
                    ),
                    "valid_hours": valid_hours,
                    "expected_hours": 24,
                    "missing_network_wet_hours": missing_wet_hours,
                    "qualifies_for_daily_statistics": qualifies,
                }
            )
    daily_sensor = pd.DataFrame(sensor_rows)
    diagnostic_lookup = diagnostics.set_index("station")[
        "systematic_review_candidate"
    ].to_dict()
    review_rows: list[dict[str, Any]] = []

    for (date, day_number), group in daily_sensor.groupby(
        ["date", "day_number"], sort=True
    ):
        totals = group.set_index("station")["daily_accumulation_mm"].dropna()
        for station, value in totals.items():
            reference = totals.drop(index=station)
            metrics = robust_distribution_metrics(reference)
            upper = metrics["q3"] + settings.daily_iqr_multiplier * metrics["iqr"]
            z_score = modified_z(value, metrics["median"], metrics["mad"])
            score = 0
            reasons: list[str] = []
            if value > upper:
                score += 1
                reasons.append("daily_above_1.5_iqr")
            if z_score > settings.daily_modified_z_review:
                score += 1
                reasons.append("daily_modified_z_above_3.5")

            neighbor_names = [
                name for name in neighbors.get(station, []) if name in totals.index
            ]
            neighbor_values = totals[neighbor_names].dropna()
            neighbor_count = len(neighbor_values)
            neighbor_median = (
                float(neighbor_values.median()) if neighbor_count else np.nan
            )
            spatial_daily = (
                neighbor_count >= settings.spatial_min_neighbors
                and value - neighbor_median >= settings.daily_spatial_difference_mm
                and safe_ratio(value, neighbor_median) >= settings.daily_spatial_ratio
            )
            if spatial_daily:
                score += 2
                reasons.append("daily_nearest_neighbor_disagreement")
            if diagnostic_lookup.get(station, False) and score > 0:
                score += 1
                reasons.append("sensor_has_systematic_review_flag")

            if score > 0:
                level = (
                    "severe_review"
                    if score >= settings.severe_review_score_threshold
                    else "review"
                )
                review_rows.append(
                    {
                        "event": event,
                        "day_number": day_number,
                        "date": date,
                        "station": station,
                        "daily_accumulation_mm": value,
                        "review_level": level,
                        "review_score": score,
                        "reasons": ";".join(reasons),
                        "network_daily_median_mm": metrics["median"],
                        "modified_z": z_score,
                        "iqr_upper_mm": upper,
                        "neighbor_daily_median_mm": neighbor_median,
                        "neighbor_count": neighbor_count,
                    }
                )

    daily_reviews = pd.DataFrame(review_rows, columns=DAILY_REVIEW_COLUMNS)
    review_keys = set(
        zip(
            daily_reviews.get("date", pd.Series(dtype="datetime64[ns]")),
            daily_reviews.get("station", pd.Series(dtype=str)),
        )
    )
    daily_sensor["daily_review_flag"] = [
        (date, station) in review_keys
        for date, station in zip(daily_sensor["date"], daily_sensor["station"])
    ]

    daily_rows: list[dict[str, Any]] = []
    for (date, day_number), group in daily_sensor.groupby(
        ["date", "day_number"], sort=True
    ):
        qualifying = group.loc[
            group["qualifies_for_daily_statistics"]
            & group["daily_accumulation_mm"].notna()
        ]
        totals = qualifying["daily_accumulation_mm"]
        review_excluded = qualifying.loc[
            ~qualifying["daily_review_flag"], "daily_accumulation_mm"
        ]
        stable_stations = qualifying["station"].tolist()
        stable_hourly = active.loc[
            date : date + pd.Timedelta(hours=23), stable_stations
        ]
        all_mean = float(totals.mean()) if not totals.empty else np.nan
        review_mean = (
            float(review_excluded.mean()) if not review_excluded.empty else np.nan
        )
        daily_rows.append(
            {
                "event": event,
                "day_number": day_number,
                "date": date,
                "mean_all_qualifying_sensors_mm": all_mean,
                "median_all_qualifying_sensors_mm": (
                    float(totals.median()) if not totals.empty else np.nan
                ),
                "trimmed_mean_10_percent_mm": trimmed_mean(
                    totals, settings.trimmed_mean_fraction
                ),
                "mean_excluding_review_sensor_days_mm": review_mean,
                "review_exclusion_difference_mm": (
                    review_mean - all_mean
                    if not pd.isna(review_mean) and not pd.isna(all_mean)
                    else np.nan
                ),
                "review_exclusion_difference_percent": (
                    100.0 * (review_mean - all_mean) / all_mean
                    if not pd.isna(review_mean) and all_mean != 0
                    else np.nan
                ),
                "qualifying_sensors": int(len(qualifying)),
                "review_sensor_days": int(qualifying["daily_review_flag"].sum()),
                "active_sensors": len(active_stations),
                "sum_of_hourly_stable_roster_means_mm": (
                    stable_hourly.mean(axis=1, skipna=True).sum(min_count=minimum_valid)
                    if stable_stations
                    else np.nan
                ),
            }
        )
    return daily_sensor, pd.DataFrame(daily_rows), daily_reviews


def calculate_hourly_network_statistics(
    event: str,
    rainfall: pd.DataFrame,
    active_stations: list[str],
    hourly_reviews: pd.DataFrame,
) -> pd.DataFrame:
    active = rainfall[active_stations]
    review_counts = (
        hourly_reviews.groupby("timestamp").size()
        if not hourly_reviews.empty
        else pd.Series(dtype=int)
    )
    severe_counts = (
        hourly_reviews.loc[hourly_reviews["review_level"] == "severe_review"]
        .groupby("timestamp")
        .size()
        if not hourly_reviews.empty
        else pd.Series(dtype=int)
    )
    frame = pd.DataFrame(
        {
            "event": event,
            "timestamp": active.index,
            "network_mean_mm": active.mean(axis=1, skipna=True).values,
            "network_median_mm": active.median(axis=1, skipna=True).values,
            "network_trimmed_mean_mm": [
                trimmed_mean(row, 0.10) for _, row in active.iterrows()
            ],
            "network_max_mm": active.max(axis=1, skipna=True).values,
            "reporting_sensors": active.count(axis=1).values,
            "active_sensors": len(active_stations),
        }
    )
    frame["review_values"] = frame["timestamp"].map(review_counts).fillna(0).astype(int)
    frame["severe_review_values"] = (
        frame["timestamp"].map(severe_counts).fillna(0).astype(int)
    )
    frame["date"] = frame["timestamp"].dt.floor("D")
    frame["hour"] = frame["timestamp"].dt.hour
    return frame


def calculate_event_summary(
    event: str,
    rainfall: pd.DataFrame,
    active_stations: list[str],
    sensor_status: pd.DataFrame,
    hourly: pd.DataFrame,
    daily: pd.DataFrame,
    invalid_rejected: pd.DataFrame,
    hourly_reviews: pd.DataFrame,
    daily_reviews: pd.DataFrame,
    diagnostics: pd.DataFrame,
    spatial_matches: pd.DataFrame,
) -> dict[str, Any]:
    active = rainfall[active_stations]
    valid_hourly = hourly.dropna(subset=["network_mean_mm"])
    peak_hour = (
        valid_hourly.loc[valid_hourly["network_mean_mm"].idxmax()]
        if not valid_hourly.empty
        else None
    )
    station_maxima = active.max(axis=0, skipna=True)
    if station_maxima.notna().any():
        max_station = str(station_maxima.idxmax())
        max_timestamp = active[max_station].idxmax()
        max_value = float(active.at[max_timestamp, max_station])
    else:
        max_station, max_timestamp, max_value = "", pd.NaT, np.nan

    wettest = (
        daily.loc[daily["mean_all_qualifying_sensors_mm"].idxmax()]
        if daily["mean_all_qualifying_sensors_mm"].notna().any()
        else None
    )
    possible = len(active) * len(active_stations)
    valid = int(active.count().sum())
    return {
        "event": event,
        "start_timestamp": active.index.min(),
        "end_timestamp": active.index.max(),
        "calendar_days": int(len(daily)),
        "input_sensors": int(
            len(sensor_status.loc[sensor_status["event"] == event])
        ),
        "active_sensors": len(active_stations),
        "excluded_sensors": int(
            (~sensor_status.loc[
                sensor_status["event"] == event, "included_in_analysis"
            ]).sum()
        ),
        "spatially_matched_active_sensors": int(
            spatial_matches.loc[
                (spatial_matches["event"] == event)
                & spatial_matches["station"].isin(active_stations)
                & spatial_matches["matched"]
            ].shape[0]
        ),
        "rejected_hourly_values": int(len(invalid_rejected)),
        "hourly_review_values": int(len(hourly_reviews)),
        "severe_hourly_review_values": int(
            (hourly_reviews["review_level"] == "severe_review").sum()
            if not hourly_reviews.empty
            else 0
        ),
        "daily_sensor_review_values": int(len(daily_reviews)),
        "systematic_sensor_review_candidates": int(
            diagnostics["systematic_review_candidate"].sum()
        ),
        "valid_cell_completeness_percent": (
            100.0 * valid / possible if possible else np.nan
        ),
        "event_total_mean_all_sensors_mm": daily[
            "mean_all_qualifying_sensors_mm"
        ].sum(min_count=len(daily)),
        "event_total_median_mm": daily[
            "median_all_qualifying_sensors_mm"
        ].sum(min_count=len(daily)),
        "event_total_trimmed_mean_mm": daily[
            "trimmed_mean_10_percent_mm"
        ].sum(min_count=len(daily)),
        "event_total_review_excluded_mean_mm": daily[
            "mean_excluding_review_sensor_days_mm"
        ].sum(min_count=len(daily)),
        "mean_daily_all_sensors_mm": daily["mean_all_qualifying_sensors_mm"].mean(),
        "wettest_day": wettest["date"] if wettest is not None else pd.NaT,
        "wettest_day_mean_mm": (
            wettest["mean_all_qualifying_sensors_mm"]
            if wettest is not None
            else np.nan
        ),
        "peak_network_mean_hour_timestamp": (
            peak_hour["timestamp"] if peak_hour is not None else pd.NaT
        ),
        "peak_network_mean_hour_mm": (
            peak_hour["network_mean_mm"] if peak_hour is not None else np.nan
        ),
        "highest_retained_station_hour_timestamp": max_timestamp,
        "highest_retained_station_hour_station": max_station,
        "highest_retained_station_hour_mm": max_value,
    }


# =============================================================================
# PLOTTING
# =============================================================================


def apply_plot_style() -> None:
    plt.rcParams.update(
        {
            "figure.dpi": 130,
            "savefig.dpi": 180,
            "font.size": 10,
            "axes.titlesize": 13,
            "axes.labelsize": 10,
            "axes.spines.top": False,
            "axes.spines.right": False,
        }
    )


def add_intensity_lines(ax: plt.Axes, data_max: float) -> None:
    top = max(3.0, data_max * 1.15 if np.isfinite(data_max) else 3.0)
    for threshold, label, color in HOURLY_INTENSITY_LINES:
        if threshold <= top:
            ax.axhline(
                threshold,
                color=color,
                linestyle="--",
                linewidth=1,
                alpha=0.85,
                label=f"{label} ({threshold:g} mm/h)",
            )


def finish_figure(fig: plt.Figure, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def legend(ax: plt.Axes, **kwargs: Any) -> None:
    handles, labels = ax.get_legend_handles_labels()
    if handles:
        ax.legend(
            frameon=True,
            facecolor="white",
            edgecolor="none",
            framealpha=0.90,
            fontsize=8,
            **kwargs,
        )


def annotate_bars(ax: plt.Axes, bars: Any) -> None:
    for bar in bars:
        height = bar.get_height()
        if np.isfinite(height):
            ax.annotate(
                f"{height:.1f}",
                (bar.get_x() + bar.get_width() / 2, height),
                xytext=(0, 4),
                textcoords="offset points",
                ha="center",
                va="bottom",
                fontsize=8,
            )


def plot_hourly_days(hourly: pd.DataFrame, root: Path) -> None:
    for (event, date), group in hourly.groupby(["event", "date"], sort=True):
        values = group.set_index("hour")["network_mean_mm"].reindex(range(24))
        review = group.set_index("hour")["review_values"].reindex(range(24)).fillna(0)
        fig, ax = plt.subplots(figsize=(10.5, 5.2))
        ax.bar(
            range(24),
            values,
            color=EVENT_COLORS.get(event, "#2878B5"),
            width=0.78,
        )
        flagged_hours = review.index[review > 0]
        if len(flagged_hours):
            ax.scatter(
                flagged_hours,
                values.loc[flagged_hours],
                color="crimson",
                marker="v",
                s=45,
                zorder=5,
                label="Hour contains retained QC-review value(s)",
            )
        add_intensity_lines(ax, float(values.max()) if values.notna().any() else 0)
        ax.set_title(
            f"{event}: Hourly Network-Mean Rainfall — {pd.Timestamp(date):%d %b %Y}"
        )
        ax.set_xlabel("Hour (PHT)")
        ax.set_ylabel("Network-mean rainfall (mm/h)")
        ax.set_xticks(range(0, 24, 2))
        ax.set_xticklabels([f"{hour:02d}:00" for hour in range(0, 24, 2)])
        ax.set_xlim(-0.7, 23.7)
        ax.set_ylim(bottom=0)
        legend(ax, ncol=2, loc="upper right")
        finish_figure(
            fig,
            root
            / "hourly_by_calendar_day"
            / f"{event}_{pd.Timestamp(date):%Y-%m-%d}_hourly.png",
        )


def plot_daily_events(daily: pd.DataFrame, root: Path) -> None:
    for event, group in daily.groupby("event", sort=False):
        group = group.sort_values("day_number")
        x = np.arange(len(group))
        values = group["mean_all_qualifying_sensors_mm"].to_numpy(float)
        labels = [
            f"Day {number}\n{date:%d %b}"
            for number, date in zip(
                group["day_number"], pd.to_datetime(group["date"])
            )
        ]
        fig, ax = plt.subplots(figsize=(7.5, 5))
        bars = ax.bar(x, values, color=EVENT_COLORS.get(event), width=0.62)
        annotate_bars(ax, bars)
        flagged = group["review_sensor_days"].to_numpy() > 0
        if flagged.any():
            ax.scatter(
                x[flagged],
                values[flagged],
                color="crimson",
                marker="v",
                s=55,
                label="Contains sensor-day review candidate(s)",
                zorder=5,
            )
        ax.set_title(f"{event}: Daily Rainfall Averaged Across Sensors")
        ax.set_xlabel("Event day")
        ax.set_ylabel("Mean daily accumulation (mm)")
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        ax.set_ylim(bottom=0)
        legend(ax, loc="upper right")
        finish_figure(fig, root / "daily_by_event" / f"{event}_daily.png")


def plot_successive_hourly(hourly: pd.DataFrame, root: Path) -> None:
    events = list(hourly["event"].drop_duplicates())
    day_map = (
        hourly[["event", "date"]]
        .drop_duplicates()
        .sort_values(["event", "date"])
        .assign(day_number=lambda frame: frame.groupby("event").cumcount() + 1)
    )
    data = hourly.merge(day_map, on=["event", "date"], how="left")
    for day_number in sorted(data["day_number"].unique()):
        subset = data.loc[data["day_number"] == day_number]
        x = np.arange(24)
        width = 0.38 if len(events) == 2 else 0.8 / len(events)
        fig, ax = plt.subplots(figsize=(11.5, 5.5))
        maxima: list[float] = []
        for event_index, event in enumerate(events):
            event_data = subset.loc[subset["event"] == event].sort_values("hour")
            values = event_data.set_index("hour")["network_mean_mm"].reindex(range(24))
            maxima.extend(values.dropna().tolist())
            offset = (event_index - (len(events) - 1) / 2) * width
            date_text = (
                f"{pd.Timestamp(event_data['date'].iloc[0]):%d %b %Y}"
                if not event_data.empty
                else "no date"
            )
            ax.bar(
                x + offset,
                values,
                width=width,
                color=EVENT_COLORS.get(event),
                label=f"{event} — {date_text}",
            )
        add_intensity_lines(ax, max(maxima) if maxima else 0)
        ax.set_title(f"Successive-Day Hourly Comparison — Day {int(day_number)}")
        ax.set_xlabel("Hour (PHT)")
        ax.set_ylabel("Network-mean rainfall (mm/h)")
        ax.set_xticks(range(0, 24, 2))
        ax.set_xticklabels([f"{hour:02d}:00" for hour in range(0, 24, 2)])
        ax.set_xlim(-0.8, 23.8)
        ax.set_ylim(bottom=0)
        legend(ax, ncol=2, loc="upper right")
        finish_figure(
            fig,
            root
            / "successive_day_hourly_comparisons"
            / f"day_{int(day_number):02d}_hourly_comparison.png",
        )


def plot_successive_daily(daily: pd.DataFrame, root: Path) -> None:
    events = list(daily["event"].drop_duplicates())
    days = sorted(daily["day_number"].unique())
    x = np.arange(len(days))
    width = 0.38 if len(events) == 2 else 0.8 / len(events)
    fig, ax = plt.subplots(figsize=(9, 5.4))
    for event_index, event in enumerate(events):
        group = daily.loc[daily["event"] == event].set_index("day_number")
        values = group["mean_all_qualifying_sensors_mm"].reindex(days)
        offset = (event_index - (len(events) - 1) / 2) * width
        bars = ax.bar(
            x + offset,
            values,
            width=width,
            color=EVENT_COLORS.get(event),
            label=event,
        )
        annotate_bars(ax, bars)
    ax.set_title("Successive-Day Daily Rainfall Comparison")
    ax.set_xlabel("Relative event day")
    ax.set_ylabel("Mean daily accumulation (mm)")
    ax.set_xticks(x)
    ax.set_xticklabels([f"Day {int(day)}" for day in days])
    ax.set_ylim(bottom=0)
    legend(ax, loc="upper right")
    finish_figure(
        fig,
        root / "successive_day_daily_comparison" / "daily_comparison.png",
    )


def plot_qc_hourly_scatter(
    cleaned: dict[str, pd.DataFrame],
    sensor_status: pd.DataFrame,
    reviews: pd.DataFrame,
    root: Path,
) -> None:
    for event, rainfall in cleaned.items():
        active_stations = sensor_status.loc[
            (sensor_status["event"] == event)
            & sensor_status["included_in_analysis"],
            "station",
        ].tolist()
        long = (
            rainfall[active_stations]
            .reset_index()
            .melt(id_vars="Date & Time", var_name="station", value_name="value_mm")
            .dropna()
        )
        long["date"] = long["Date & Time"].dt.floor("D")
        long["hour"] = long["Date & Time"].dt.hour
        event_reviews = reviews.loc[reviews["event"] == event]
        for date, group in long.groupby("date", sort=True):
            fig, ax = plt.subplots(figsize=(10.5, 5.2))
            ax.scatter(
                group["hour"],
                group["value_mm"],
                color="#777777",
                alpha=0.45,
                s=18,
                label="Retained station values",
            )
            flagged = event_reviews.loc[
                event_reviews["timestamp"].dt.floor("D") == date
            ]
            if not flagged.empty:
                ax.scatter(
                    flagged["timestamp"].dt.hour,
                    flagged["value_mm"],
                    color="crimson",
                    marker="x",
                    s=65,
                    linewidth=1.5,
                    label="Retained QC-review values",
                    zorder=5,
                )
            ax.set_title(
                f"{event}: Hourly Station QC Distribution — {pd.Timestamp(date):%d %b %Y}"
            )
            ax.set_xlabel("Hour (PHT)")
            ax.set_ylabel("Station rainfall (mm/h)")
            ax.set_xticks(range(0, 24, 2))
            ax.set_xticklabels([f"{hour:02d}:00" for hour in range(0, 24, 2)])
            ax.set_xlim(-0.7, 23.7)
            ax.set_ylim(bottom=0)
            legend(ax, loc="upper right")
            finish_figure(
                fig,
                root
                / "QC_hourly_station_distributions"
                / f"{event}_{pd.Timestamp(date):%Y-%m-%d}_qc.png",
            )


def plot_sensor_day_bars(
    daily_sensor: pd.DataFrame,
    root: Path,
) -> None:
    for (event, date), group in daily_sensor.groupby(["event", "date"], sort=True):
        group = group.loc[group["daily_accumulation_mm"].notna()].sort_values(
            "daily_accumulation_mm"
        )
        if group.empty:
            continue
        colors = np.where(group["daily_review_flag"], "crimson", "#4C90C0")
        fig_height = max(5.5, 0.24 * len(group))
        fig, ax = plt.subplots(figsize=(9, fig_height))
        ax.barh(group["station"], group["daily_accumulation_mm"], color=colors)
        ax.set_title(
            f"{event}: Sensor Daily Accumulations — {pd.Timestamp(date):%d %b %Y}"
        )
        ax.set_xlabel("Daily accumulation (mm)")
        ax.set_ylabel("")
        finish_figure(
            fig,
            root
            / "QC_sensor_day_accumulations"
            / f"{event}_{pd.Timestamp(date):%Y-%m-%d}_sensor_totals.png",
        )


def plot_sensitivity(daily: pd.DataFrame, root: Path) -> None:
    metrics = [
        ("mean_all_qualifying_sensors_mm", "Mean"),
        ("median_all_qualifying_sensors_mm", "Median"),
        ("trimmed_mean_10_percent_mm", "10% trimmed mean"),
        ("mean_excluding_review_sensor_days_mm", "Review-excluded mean"),
    ]
    for event, group in daily.groupby("event", sort=False):
        group = group.sort_values("day_number")
        x = np.arange(len(group))
        width = 0.19
        fig, ax = plt.subplots(figsize=(10, 5.4))
        for metric_index, (column, label) in enumerate(metrics):
            offset = (metric_index - 1.5) * width
            ax.bar(x + offset, group[column], width=width, label=label)
        ax.set_title(f"{event}: Daily Rainfall Sensitivity Statistics")
        ax.set_xlabel("Relative event day")
        ax.set_ylabel("Daily accumulation statistic (mm)")
        ax.set_xticks(x)
        ax.set_xticklabels([f"Day {int(day)}" for day in group["day_number"]])
        ax.set_ylim(bottom=0)
        legend(ax, ncol=2, loc="upper right")
        finish_figure(
            fig, root / "QC_sensitivity" / f"{event}_daily_sensitivity.png"
        )


def plot_spatial_review_map(
    spatial_matches: pd.DataFrame,
    hourly_reviews: pd.DataFrame,
    daily_reviews: pd.DataFrame,
    sensor_status: pd.DataFrame,
    root: Path,
) -> None:
    for event, matches in spatial_matches.groupby("event", sort=False):
        active = sensor_status.loc[
            (sensor_status["event"] == event)
            & sensor_status["included_in_analysis"],
            "station",
        ]
        matches = matches.loc[matches["matched"] & matches["station"].isin(active)]
        if matches.empty:
            continue
        counts = pd.concat(
            [
                hourly_reviews.loc[hourly_reviews["event"] == event, ["station"]],
                daily_reviews.loc[daily_reviews["event"] == event, ["station"]],
            ],
            ignore_index=True,
        ).value_counts("station")
        matches = matches.copy()
        matches["review_count"] = matches["station"].map(counts).fillna(0)
        flagged = matches["review_count"] > 0
        fig, ax = plt.subplots(figsize=(7.5, 7.5))
        ax.scatter(
            matches.loc[~flagged, "x_m"],
            matches.loc[~flagged, "y_m"],
            color="#777777",
            s=24,
            alpha=0.65,
            label="Matched active gauge",
        )
        if flagged.any():
            ax.scatter(
                matches.loc[flagged, "x_m"],
                matches.loc[flagged, "y_m"],
                color="crimson",
                s=60 + 15 * matches.loc[flagged, "review_count"],
                label="Gauge with review flag(s)",
                zorder=5,
            )
            for _, row in matches.loc[flagged].iterrows():
                ax.annotate(
                    row["station"],
                    (row["x_m"], row["y_m"]),
                    xytext=(4, 4),
                    textcoords="offset points",
                    fontsize=7,
                )
        ax.set_title(f"{event}: Spatial Distribution of QC Review Flags")
        ax.set_xlabel("UTM Zone 51N Easting (m)")
        ax.set_ylabel("UTM Zone 51N Northing (m)")
        ax.set_aspect("equal", adjustable="datalim")
        legend(ax, loc="upper right")
        finish_figure(fig, root / "QC_spatial" / f"{event}_spatial_review.png")


# =============================================================================
# TEXT OUTPUTS
# =============================================================================


def fmt_mm(value: Any) -> str:
    return "not available" if pd.isna(value) else f"{float(value):.2f} mm"


def fmt_timestamp(value: Any) -> str:
    return (
        "not available"
        if pd.isna(value)
        else pd.Timestamp(value).strftime("%d %b %Y %H:%M PHT")
    )


def write_summary(
    path: Path,
    summaries: pd.DataFrame,
    daily: pd.DataFrame,
    sensor_status: pd.DataFrame,
    diagnostics: pd.DataFrame,
    settings: QCSettings,
    gauge_name_field: str,
) -> None:
    lines = [
        "CDEHab vs MDHab Rainfall Event Comparison — QC Review Version",
        "================================================================",
        "",
        "Method",
        "------",
        (
            "Only invalid observations, manual rejections, and sensors identified as "
            "unusable are omitted. Statistical rainfall extremes are retained and "
            "reported as review candidates."
        ),
        (
            f"Spatial verification uses up to {settings.spatial_max_neighbors} nearest "
            f"gauges, primarily within {settings.spatial_radius_km:g} km. The detected "
            f"shapefile name field was '{gauge_name_field}'."
        ),
        (
            "Daily values are accumulated per qualifying sensor first. A sensor-day "
            f"requires at least {settings.minimum_daily_completeness:.0%} completeness "
            "and no missing hour during network-wet conditions."
        ),
        "",
        "Event statistics",
        "----------------",
    ]
    for _, row in summaries.iterrows():
        event = row["event"]
        exclusions = sensor_status.loc[
            (sensor_status["event"] == event)
            & (~sensor_status["included_in_analysis"]),
            "status",
        ].value_counts()
        exclusion_text = (
            ", ".join(f"{name}: {count}" for name, count in exclusions.items())
            if not exclusions.empty
            else "none"
        )
        candidates = diagnostics.loc[
            (diagnostics["event"] == event)
            & diagnostics["systematic_review_candidate"],
            ["station", "review_reasons"],
        ]
        lines.extend(
            [
                event,
                f"  Period: {fmt_timestamp(row['start_timestamp'])} to {fmt_timestamp(row['end_timestamp'])}",
                (
                    f"  Sensors: {int(row['active_sensors'])} active of "
                    f"{int(row['input_sensors'])}; excluded — {exclusion_text}"
                ),
                (
                    f"  Spatial matches: {int(row['spatially_matched_active_sensors'])} "
                    f"of {int(row['active_sensors'])} active sensors"
                ),
                f"  Hard-rejected hourly values: {int(row['rejected_hourly_values'])}",
                (
                    f"  Retained hourly review values: {int(row['hourly_review_values'])}, "
                    f"including {int(row['severe_hourly_review_values'])} severe-review value(s)"
                ),
                f"  Sensor-day review candidates: {int(row['daily_sensor_review_values'])}",
                (
                    "  Event-total daily means: "
                    f"ordinary {fmt_mm(row['event_total_mean_all_sensors_mm'])}; "
                    f"median-based {fmt_mm(row['event_total_median_mm'])}; "
                    f"10% trimmed {fmt_mm(row['event_total_trimmed_mean_mm'])}; "
                    f"review-excluded sensitivity {fmt_mm(row['event_total_review_excluded_mean_mm'])}"
                ),
                (
                    f"  Wettest day: {pd.Timestamp(row['wettest_day']):%d %b %Y} "
                    f"({fmt_mm(row['wettest_day_mean_mm'])})"
                ),
                (
                    f"  Peak network-mean hour: "
                    f"{fmt_timestamp(row['peak_network_mean_hour_timestamp'])} "
                    f"({fmt_mm(row['peak_network_mean_hour_mm'])})"
                ),
                (
                    "  Highest retained station-hour: "
                    f"{row['highest_retained_station_hour_station']}, "
                    f"{fmt_timestamp(row['highest_retained_station_hour_timestamp'])} "
                    f"({fmt_mm(row['highest_retained_station_hour_mm'])})"
                ),
                "  Systematic sensor-review candidates:",
            ]
        )
        if candidates.empty:
            lines.append("    none")
        else:
            for _, candidate in candidates.iterrows():
                lines.append(
                    f"    {candidate['station']}: {candidate['review_reasons']}"
                )
        lines.append("")

    lines.extend(
        [
            "Interpretation",
            "--------------",
            (
                "The ordinary mean is the primary continuity statistic. The median, "
                "trimmed mean, and review-excluded mean are sensitivity estimates; "
                "they do not constitute automatic replacements for the observed data."
            ),
            (
                "Review hourly_values_for_review.csv, sensor_day_outliers.csv, "
                "sensor_behavior_diagnostics.csv, and station_spatial_matches.csv "
                "before approving exclusions."
            ),
            "",
            "Daily sensitivity table",
            "-----------------------",
        ]
    )
    for _, row in daily.iterrows():
        lines.append(
            f"{row['event']} Day {int(row['day_number'])} ({pd.Timestamp(row['date']):%d %b %Y}): "
            f"mean {fmt_mm(row['mean_all_qualifying_sensors_mm'])}; "
            f"median {fmt_mm(row['median_all_qualifying_sensors_mm'])}; "
            f"trimmed {fmt_mm(row['trimmed_mean_10_percent_mm'])}; "
            f"review-excluded {fmt_mm(row['mean_excluding_review_sensor_days_mm'])}"
        )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_qc_readme(path: Path) -> None:
    text = """QC OUTPUT GUIDE
===============

Invalid values are rejected; statistical extremes are retained for review.

Key tables
----------
invalid_values_rejected.csv
    Values actually omitted from calculations.

hourly_values_for_review.csv
    Retained hourly anomalies and every component of their review score.

sensor_day_outliers.csv
    Retained daily sensor accumulations that require review.

sensor_behavior_diagnostics.csv
    Persistent bias, correlation, zero-reporting, and lag diagnostics.

station_spatial_matches.csv
    Dataset-to-shapefile matching. Correct MANUAL_STATION_ALIASES in the script
    if any match is wrong or missing.

station_neighbors.csv
    Nearest gauges used for spatial verification and their distances.

daily_statistics_sensitivity.csv
    Ordinary mean, median, trimmed mean, and review-excluded sensitivity mean.

manual_qc_overrides.csv
    Optional actions: reject_hour, retain_hour, exclude_sensor, include_sensor.
    Use PHT timestamps such as 2025-07-21 12:00. Rerun the script after editing.
"""
    path.write_text(text, encoding="utf-8")


# =============================================================================
# OUTPUT ORCHESTRATION
# =============================================================================


def write_outputs(
    output_directory: Path,
    cleaned: dict[str, pd.DataFrame],
    hourly: pd.DataFrame,
    daily_sensor: pd.DataFrame,
    daily: pd.DataFrame,
    summaries: pd.DataFrame,
    invalid_rejected: pd.DataFrame,
    hourly_reviews: pd.DataFrame,
    daily_reviews: pd.DataFrame,
    sensor_status: pd.DataFrame,
    diagnostics: pd.DataFrame,
    spatial_matches: pd.DataFrame,
    neighbors: pd.DataFrame,
    manual_actions: pd.DataFrame,
    data_issues: pd.DataFrame,
    settings: QCSettings,
    gauge_name_field: str,
) -> None:
    output_directory.mkdir(parents=True, exist_ok=True)
    tables = output_directory / "Tables"
    plots = output_directory / "Plots"

    for event, frame in cleaned.items():
        output = frame.reset_index()
        save_csv(output, tables / f"cleaned_hourly_{event}.csv")
    save_csv(invalid_rejected, tables / "invalid_values_rejected.csv")
    save_csv(hourly_reviews, tables / "hourly_values_for_review.csv")
    save_csv(daily_reviews, tables / "sensor_day_outliers.csv")
    save_csv(diagnostics, tables / "sensor_behavior_diagnostics.csv")
    save_csv(sensor_status, tables / "sensor_status.csv")
    save_csv(spatial_matches, tables / "station_spatial_matches.csv")
    save_csv(neighbors, tables / "station_neighbors.csv")
    save_csv(hourly, tables / "hourly_network_statistics.csv")
    save_csv(daily_sensor, tables / "daily_station_accumulations.csv")
    save_csv(daily, tables / "daily_statistics_sensitivity.csv")
    save_csv(summaries, tables / "event_summary_statistics.csv")
    save_csv(manual_actions, tables / "manual_override_actions_applied.csv")
    save_csv(data_issues, tables / "data_issues.csv")

    override_path = output_directory / "manual_qc_overrides.csv"
    if not override_path.exists():
        save_csv(pd.DataFrame(columns=MANUAL_OVERRIDE_COLUMNS), override_path)
    write_qc_readme(output_directory / "QC_README.txt")
    write_summary(
        output_directory / "CDEHab_vs_MDHab_summary.txt",
        summaries,
        daily,
        sensor_status,
        diagnostics,
        settings,
        gauge_name_field,
    )

    apply_plot_style()
    plot_hourly_days(hourly, plots)
    plot_daily_events(daily, plots)
    plot_successive_hourly(hourly, plots)
    plot_successive_daily(daily, plots)
    plot_qc_hourly_scatter(cleaned, sensor_status, hourly_reviews, plots)
    plot_sensor_day_bars(daily_sensor, plots)
    plot_sensitivity(daily, plots)
    plot_spatial_review_map(
        spatial_matches, hourly_reviews, daily_reviews, sensor_status, plots
    )


# =============================================================================
# MAIN ANALYSIS
# =============================================================================


def run_analysis(
    event_files: list[tuple[str, Path]],
    shapefile_path: Path,
    output_directory: Path,
    settings: QCSettings,
    gauge_name_field: str | None = None,
    disable_spatial: bool = False,
    gauge_locations_override: pd.DataFrame | None = None,
) -> None:
    overrides = load_manual_overrides(output_directory / "manual_qc_overrides.csv")
    loaded: dict[str, pd.DataFrame] = {}
    statuses: list[pd.DataFrame] = []
    invalid_rows: list[dict[str, Any]] = []
    issues: list[dict[str, str]] = []
    manual_retained: dict[str, set[tuple[pd.Timestamp, str]]] = {}
    manual_actions: list[pd.DataFrame] = []

    for event, file_path in event_files:
        print(f"Reading {event}: {file_path}")
        rainfall, rejected, event_issues = load_event(event, file_path, settings)
        status = identify_sensor_status(event, rainfall, settings, overrides)
        rainfall, manual_rejected, retained, action_frame = apply_manual_hour_actions(
            event, rainfall, overrides
        )
        loaded[event] = rainfall
        statuses.append(status)
        invalid_rows.extend(rejected + manual_rejected)
        issues.extend(event_issues)
        manual_retained[event] = retained
        if not action_frame.empty:
            manual_actions.append(action_frame)

    sensor_status = pd.concat(statuses, ignore_index=True)
    events_and_stations = {
        event: sensor_status.loc[
            (sensor_status["event"] == event)
            & sensor_status["included_in_analysis"],
            "station",
        ].tolist()
        for event, _ in event_files
    }
    all_sensor_names = sorted(
        {station for stations in events_and_stations.values() for station in stations}
    )

    if disable_spatial:
        gauge_locations = pd.DataFrame(columns=["gauge_index", "gauge_name", "x_m", "y_m"])
        detected_name_field = "spatial_verification_disabled"
    elif gauge_locations_override is not None:
        gauge_locations = gauge_locations_override.copy()
        detected_name_field = "test_or_programmatic_override"
    else:
        print(f"Reading gauge locations: {shapefile_path}")
        gauge_locations, detected_name_field = load_gauge_locations(
            shapefile_path, all_sensor_names, gauge_name_field
        )

    if disable_spatial:
        spatial_matches = pd.DataFrame(
            [
                {
                    "event": event,
                    "station": station,
                    "alias_used": MANUAL_STATION_ALIASES.get(station, station),
                    "matched": False,
                    "matched_gauge_name": "",
                    "match_score": np.nan,
                    "match_method": "spatial_disabled",
                    "x_m": np.nan,
                    "y_m": np.nan,
                }
                for event, stations in events_and_stations.items()
                for station in stations
            ]
        )
    else:
        spatial_matches = match_stations_to_gauges(
            events_and_stations, gauge_locations, settings
        )
    neighbor_table = build_neighbor_table(spatial_matches, sensor_status, settings)
    neighbors_by_event = neighbor_map_from_table(neighbor_table)

    cleaned_for_analysis: dict[str, pd.DataFrame] = {}
    all_diagnostics: list[pd.DataFrame] = []
    all_hourly_reviews: list[pd.DataFrame] = []
    all_daily_reviews: list[pd.DataFrame] = []
    all_daily_sensor: list[pd.DataFrame] = []
    all_daily: list[pd.DataFrame] = []
    all_hourly: list[pd.DataFrame] = []
    summary_rows: list[dict[str, Any]] = []

    invalid_frame = pd.DataFrame(invalid_rows, columns=INVALID_COLUMNS)
    for event, _ in event_files:
        status = sensor_status.loc[sensor_status["event"] == event]
        active_stations = status.loc[
            status["included_in_analysis"], "station"
        ].tolist()
        rainfall = loaded[event].copy()
        inactive = [station for station in rainfall if station not in active_stations]
        if inactive:
            rainfall.loc[:, inactive] = np.nan
        cleaned_for_analysis[event] = rainfall
        event_neighbors = neighbors_by_event.get(event, {})

        diagnostics = calculate_sensor_diagnostics(
            event, rainfall, active_stations, event_neighbors, settings
        )
        hourly_reviews = score_hourly_reviews(
            event,
            rainfall,
            active_stations,
            event_neighbors,
            diagnostics,
            manual_retained[event],
            settings,
        )
        daily_sensor, daily, daily_reviews = calculate_daily_statistics(
            event,
            rainfall,
            active_stations,
            event_neighbors,
            diagnostics,
            settings,
        )
        hourly = calculate_hourly_network_statistics(
            event, rainfall, active_stations, hourly_reviews
        )
        event_invalid = invalid_frame.loc[invalid_frame["event"] == event]
        summary = calculate_event_summary(
            event,
            rainfall,
            active_stations,
            sensor_status,
            hourly,
            daily,
            event_invalid,
            hourly_reviews,
            daily_reviews,
            diagnostics,
            spatial_matches,
        )

        all_diagnostics.append(diagnostics)
        all_hourly_reviews.append(hourly_reviews)
        all_daily_reviews.append(daily_reviews)
        all_daily_sensor.append(daily_sensor)
        all_daily.append(daily)
        all_hourly.append(hourly)
        summary_rows.append(summary)

    diagnostics = pd.concat(all_diagnostics, ignore_index=True)
    hourly_reviews = pd.concat(all_hourly_reviews, ignore_index=True)
    daily_reviews = pd.concat(all_daily_reviews, ignore_index=True)
    daily_sensor = pd.concat(all_daily_sensor, ignore_index=True)
    daily = pd.concat(all_daily, ignore_index=True)
    hourly = pd.concat(all_hourly, ignore_index=True)
    summaries = pd.DataFrame(summary_rows)
    manual_action_frame = (
        pd.concat(manual_actions, ignore_index=True)
        if manual_actions
        else pd.DataFrame(
            columns=["event", "timestamp", "station", "action", "reason", "applied"]
        )
    )
    data_issues = pd.DataFrame(issues, columns=DATA_ISSUE_COLUMNS)

    write_outputs(
        output_directory,
        cleaned_for_analysis,
        hourly,
        daily_sensor,
        daily,
        summaries,
        invalid_frame,
        hourly_reviews,
        daily_reviews,
        sensor_status,
        diagnostics,
        spatial_matches,
        neighbor_table,
        manual_action_frame,
        data_issues,
        settings,
        detected_name_field,
    )
    print(f"\nAnalysis complete. Outputs written to:\n{output_directory}")


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="QC and compare the CDEHab and MDHab rainfall events."
    )
    parser.add_argument("--cdehab", type=Path, default=CDEHAB_FILE)
    parser.add_argument("--mdhab", type=Path, default=MDHAB_FILE)
    parser.add_argument("--shapefile", type=Path, default=RAIN_GAUGE_SHAPEFILE)
    parser.add_argument("--output", type=Path, default=OUTPUT_DIRECTORY)
    parser.add_argument(
        "--gauge-name-field",
        default=None,
        help="Explicit shapefile field containing gauge names; otherwise auto-detected.",
    )
    parser.add_argument(
        "--no-spatial",
        action="store_true",
        help="Run without shapefile verification if GIS dependencies are unavailable.",
    )
    return parser


def main() -> int:
    args = build_argument_parser().parse_args()
    try:
        run_analysis(
            [("CDEHab", args.cdehab), ("MDHab", args.mdhab)],
            args.shapefile,
            args.output,
            DEFAULT_SETTINGS,
            gauge_name_field=args.gauge_name_field,
            disable_spatial=args.no_spatial,
        )
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
