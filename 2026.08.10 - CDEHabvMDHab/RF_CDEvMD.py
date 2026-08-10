"""
Compare the CDEHab (19-21 July 2025) and MDHab (7-9 August 2026)
hourly rainfall events.

The script:
  1. reads both Excel workbooks, including decimal-comma rainfall values;
  2. identifies unusable sensors and conservative cell-level outliers;
  3. omits flagged values from all accumulation and averaging calculations;
  4. calculates hourly network means and daily station-first accumulations;
  5. produces individual-day, event-daily, and successive-day comparison plots;
  6. writes auditable CSV outputs and a plain-language comparison summary.

Install the required packages, if necessary:
    py -m pip install pandas numpy matplotlib openpyxl

Run with the configured Windows paths:
    py rainfall_event_comparison.py

Optional path overrides:
    py rainfall_event_comparison.py ^
      --cdehab "C:\\path\\to\\CDEHab.xlsx" ^
      --mdhab "C:\\path\\to\\MDHab.xlsx" ^
      --output "C:\\path\\to\\Outputs"

Important methodological choices
--------------------------------
* Rainfall outlier screening is one-sided. Legitimate zero rainfall is not
  treated as an outlier.
* A statistically high value is omitted only if it is extreme both relative
  to other active sensors at the same hour and relative to the same sensor's
  surrounding hours. The conservative thresholds are editable below.
* Sensors that are entirely missing, entirely zero, or constant for the whole
  event are omitted by default and documented in sensor_status.csv.
* The primary daily value is calculated station-first: hourly rainfall is
  accumulated for each qualifying sensor-day, then those daily totals are
  averaged across sensors. A sensor-day must retain at least 90% of its 24
  expected hourly observations. Missing/outlier hours are not imputed.
* Because the two events have different station rosters, event comparisons use
  network means. They are not paired comparisons of identical stations.
"""

from __future__ import annotations

import argparse
import math
import re
import sys
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np
    import pandas as pd
except ImportError as exc:  # pragma: no cover - provides a clearer user message
    raise SystemExit(
        "A required package is missing. Run: "
        "py -m pip install pandas numpy matplotlib openpyxl"
    ) from exc


# ---------------------------------------------------------------------------
# USER CONFIGURATION
# ---------------------------------------------------------------------------

CDEHAB_FILE = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.10 - CDEHabvMDHab"
    r"\Processed Data\CDEHab.xlsx"
)

MDHAB_FILE = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.10 - CDEHabvMDHab"
    r"\Processed Data\MDHab.xlsx"
)

OUTPUT_DIRECTORY = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.10 - CDEHabvMDHab\Outputs"
)


@dataclass(frozen=True)
class OutlierSettings:
    """Editable cleaning thresholds."""

    # Values outside the physically admissible input range are always omitted.
    hard_max_hourly_mm: float = 300.0

    # Statistical screening is deliberately conservative.
    statistical_screening: bool = True
    minimum_statistical_outlier_mm: float = 75.0
    cross_section_mad_multiplier: float = 8.0
    cross_section_iqr_multiplier: float = 4.0
    cross_section_ratio_multiplier: float = 6.0
    temporal_mad_multiplier: float = 8.0
    temporal_ratio_multiplier: float = 6.0
    temporal_window_hours: int = 2
    minimum_reporting_sensors_for_screening: int = 8

    # Sensor and daily-completeness rules.
    exclude_all_zero_sensors: bool = True
    minimum_daily_completeness: float = 0.90


DEFAULT_SETTINGS = OutlierSettings()

EVENT_COLORS = {
    "CDEHab": "#2878B5",
    "MDHab": "#E07A1F",
}

HOURLY_INTENSITY_LINES = (
    (2.5, "Moderate", "#F39C12"),
    (7.5, "Heavy", "#D62728"),
    (15.0, "Intense", "#8B0000"),
    (30.0, "Torrential", "#7B2CBF"),
)

OUTLIER_AUDIT_COLUMNS = [
    "event",
    "timestamp",
    "station",
    "original_value",
    "reason",
    "network_median_mm",
    "cross_section_threshold_mm",
    "temporal_threshold_mm",
]

DATA_ISSUE_COLUMNS = ["event", "issue", "details"]


# ---------------------------------------------------------------------------
# INPUT AND CLEANING
# ---------------------------------------------------------------------------


def parse_datetime_cell(value: Any) -> pd.Timestamp:
    """Parse Excel datetimes and text such as '19 07 2025 00:00'."""

    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, datetime):
        return pd.Timestamp(value)

    # Excel serial date/time values use 1899-12-30 as the practical origin.
    if isinstance(value, (int, float, np.integer, np.floating)):
        try:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(value), unit="D")
        except (ValueError, TypeError, OverflowError):
            return pd.NaT

    text = re.sub(r"\s+", " ", str(value).strip())
    formats = (
        "%d %m %Y %H:%M",
        "%d/%m/%Y %H:%M",
        "%d-%m-%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    )
    for date_format in formats:
        try:
            return pd.Timestamp(datetime.strptime(text, date_format))
        except ValueError:
            continue

    return pd.to_datetime(text, dayfirst=True, errors="coerce")


def parse_rainfall_cell(value: Any) -> float:
    """Convert ordinary numbers and decimal-comma text to millimetres."""

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
    """Return True when a failed numeric value is not just a blank marker."""

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
    candidates = (
        "date & time",
        "date and time",
        "date time",
        "datetime",
        "date pht",
        "date",
    )
    for candidate in candidates:
        if candidate in normalized:
            return normalized[candidate]
    return columns[0]


def load_event(
    event: str,
    file_path: Path,
    settings: OutlierSettings,
) -> tuple[pd.DataFrame, list[dict[str, Any]], list[dict[str, str]]]:
    """Read and minimally clean one workbook before statistical screening."""

    if not file_path.exists():
        raise FileNotFoundError(
            f"Input workbook not found for {event}:\n{file_path}\n\n"
            "Check the path in USER CONFIGURATION or use the command-line "
            "path overrides."
        )

    try:
        raw = pd.read_excel(file_path, sheet_name=0, dtype=object, engine="openpyxl")
    except ImportError as exc:
        raise RuntimeError(
            "Reading .xlsx files requires openpyxl. Run: "
            "py -m pip install openpyxl"
        ) from exc

    raw = raw.dropna(axis=1, how="all")
    if raw.empty or raw.shape[1] < 2:
        raise ValueError(f"{file_path.name} does not contain a usable rainfall table.")

    datetime_column = find_datetime_column(list(raw.columns))
    timestamps = raw[datetime_column].map(parse_datetime_cell)
    invalid_timestamp_count = int(timestamps.isna().sum())
    data_issues: list[dict[str, str]] = []
    if invalid_timestamp_count:
        data_issues.append(
            {
                "event": event,
                "issue": "invalid_timestamp_rows_removed",
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
    outlier_audit: list[dict[str, Any]] = []

    for station in station_columns:
        parsed = original[station].map(parse_rainfall_cell)
        numeric[str(station).strip()] = parsed

        malformed = parsed.isna() & original[station].map(is_meaningful_unparsed_value)
        for row_index in original.index[malformed]:
            outlier_audit.append(
                {
                    "event": event,
                    "timestamp": timestamps.loc[row_index],
                    "station": str(station).strip(),
                    "original_value": original.at[row_index, station],
                    "reason": "non_numeric_value",
                    "network_median_mm": np.nan,
                    "cross_section_threshold_mm": np.nan,
                    "temporal_threshold_mm": np.nan,
                }
            )

    numeric.index = pd.DatetimeIndex(timestamps.values, name="Date & Time")

    # Flag invalid/hard-limit values before resolving duplicate timestamps.
    for timestamp, row in numeric.iterrows():
        for station, value in row.items():
            if pd.isna(value):
                continue
            reason = None
            if value < 0:
                reason = "negative_rainfall"
            elif value > settings.hard_max_hourly_mm:
                reason = f"above_hard_limit_{settings.hard_max_hourly_mm:g}_mm"
            if reason is not None:
                outlier_audit.append(
                    {
                        "event": event,
                        "timestamp": timestamp,
                        "station": station,
                        "original_value": value,
                        "reason": reason,
                        "network_median_mm": np.nan,
                        "cross_section_threshold_mm": np.nan,
                        "temporal_threshold_mm": np.nan,
                    }
                )
                numeric.at[timestamp, station] = np.nan

    duplicate_rows = int(numeric.index.duplicated(keep=False).sum())
    if duplicate_rows:
        data_issues.append(
            {
                "event": event,
                "issue": "duplicate_timestamp_rows_averaged",
                "details": str(duplicate_rows),
            }
        )
        numeric = numeric.groupby(level=0).mean()

    numeric = numeric.sort_index()
    start = numeric.index.min().floor("D")
    end = numeric.index.max().floor("D") + pd.Timedelta(hours=23)
    full_index = pd.date_range(start, end, freq="h", name="Date & Time")
    missing_timestamp_count = int(len(full_index.difference(numeric.index)))
    if missing_timestamp_count:
        data_issues.append(
            {
                "event": event,
                "issue": "missing_hourly_timestamps_inserted",
                "details": str(missing_timestamp_count),
            }
        )
    numeric = numeric.reindex(full_index)
    return numeric, outlier_audit, data_issues


def identify_sensor_status(
    event: str,
    rainfall: pd.DataFrame,
    settings: OutlierSettings,
) -> pd.DataFrame:
    """Identify sensors unsuitable for event-level comparisons."""

    rows: list[dict[str, Any]] = []
    expected = len(rainfall)

    for station in rainfall.columns:
        series = rainfall[station]
        valid = series.dropna()
        positive_count = int((valid > 0).sum())
        included = True
        status = "active"

        if valid.empty:
            status = "all_missing"
            included = False
        elif positive_count == 0 and settings.exclude_all_zero_sensors:
            status = "all_zero"
            included = False
        elif valid.nunique(dropna=True) == 1 and len(valid) >= math.ceil(0.90 * expected):
            # An exactly constant positive value across essentially a whole
            # multi-day event is characteristic of a stuck sensor.
            status = "constant_value"
            included = False
        elif positive_count == 0:
            status = "active_all_zero"

        rows.append(
            {
                "event": event,
                "station": station,
                "status": status,
                "included_in_analysis": included,
                "valid_observations": int(valid.count()),
                "expected_observations": expected,
                "positive_observations": positive_count,
                "minimum_mm": valid.min() if not valid.empty else np.nan,
                "maximum_mm": valid.max() if not valid.empty else np.nan,
            }
        )

    return pd.DataFrame(rows)


def robust_upper_threshold(
    values: pd.Series,
    minimum_mm: float,
    mad_multiplier: float,
    iqr_multiplier: float,
    ratio_multiplier: float,
) -> tuple[float, float]:
    """Return (median, conservative upper threshold) for positive rainfall."""

    values = values.dropna().astype(float)
    median = float(values.median())
    mad = float((values - median).abs().median())
    robust_sigma = 1.4826 * mad
    q1 = float(values.quantile(0.25))
    q3 = float(values.quantile(0.75))
    iqr = q3 - q1

    return median, max(
        minimum_mm,
        median + mad_multiplier * robust_sigma,
        q3 + iqr_multiplier * iqr,
        (median + 0.5) * ratio_multiplier,
    )


def temporal_upper_threshold(
    values: pd.Series,
    settings: OutlierSettings,
) -> float:
    values = values.dropna().astype(float)
    if values.empty:
        return np.inf

    median = float(values.median())
    mad = float((values - median).abs().median())
    robust_sigma = 1.4826 * mad
    q1 = float(values.quantile(0.25))
    q3 = float(values.quantile(0.75))
    iqr = q3 - q1

    return max(
        settings.minimum_statistical_outlier_mm,
        median + settings.temporal_mad_multiplier * robust_sigma,
        q3 + settings.cross_section_iqr_multiplier * iqr,
        (median + 0.5) * settings.temporal_ratio_multiplier,
    )


def omit_statistical_outliers(
    event: str,
    rainfall: pd.DataFrame,
    active_stations: list[str],
    settings: OutlierSettings,
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    """Apply conservative cross-sectional + temporal high-outlier screening."""

    cleaned = rainfall.copy()
    inactive = [column for column in cleaned.columns if column not in active_stations]
    if inactive:
        cleaned.loc[:, inactive] = np.nan

    if not settings.statistical_screening or not active_stations:
        return cleaned, []

    candidates: list[tuple[pd.Timestamp, str, float, float, float, float]] = []

    # Calculate all candidates before removing any so results do not depend on
    # the order in which cells are processed.
    for timestamp, row in cleaned[active_stations].iterrows():
        # Include valid zeros when establishing the network distribution. This
        # allows the screen to catch an isolated large spike during an otherwise
        # dry hour. The 75 mm/h floor keeps small legitimate local differences
        # from being classified as outliers when the network median is zero.
        reported = row.dropna()
        if len(reported) < settings.minimum_reporting_sensors_for_screening:
            continue

        network_median, cross_threshold = robust_upper_threshold(
            reported,
            settings.minimum_statistical_outlier_mm,
            settings.cross_section_mad_multiplier,
            settings.cross_section_iqr_multiplier,
            settings.cross_section_ratio_multiplier,
        )

        for station, value in reported[reported > cross_threshold].items():
            window = cleaned.loc[
                timestamp - pd.Timedelta(hours=settings.temporal_window_hours) :
                timestamp + pd.Timedelta(hours=settings.temporal_window_hours),
                station,
            ]
            window = window.loc[window.index != timestamp]
            temporal_threshold = temporal_upper_threshold(window, settings)
            if value > temporal_threshold:
                candidates.append(
                    (
                        timestamp,
                        station,
                        float(value),
                        network_median,
                        cross_threshold,
                        temporal_threshold,
                    )
                )

    audit: list[dict[str, Any]] = []
    for (
        timestamp,
        station,
        value,
        network_median,
        cross_threshold,
        temporal_threshold,
    ) in candidates:
        cleaned.at[timestamp, station] = np.nan
        audit.append(
            {
                "event": event,
                "timestamp": timestamp,
                "station": station,
                "original_value": value,
                "reason": "statistical_high_outlier_spatial_and_temporal",
                "network_median_mm": network_median,
                "cross_section_threshold_mm": cross_threshold,
                "temporal_threshold_mm": temporal_threshold,
            }
        )

    return cleaned, audit


# ---------------------------------------------------------------------------
# AGGREGATION
# ---------------------------------------------------------------------------


def calculate_event_statistics(
    event: str,
    cleaned: pd.DataFrame,
    sensor_status: pd.DataFrame,
    audit: list[dict[str, Any]],
    settings: OutlierSettings,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Calculate hourly means, daily sensor totals, and event statistics."""

    active_stations = sensor_status.loc[
        sensor_status["included_in_analysis"], "station"
    ].tolist()
    if not active_stations:
        raise ValueError(f"No active sensors remain for {event} after cleaning.")

    active = cleaned[active_stations]
    hourly = pd.DataFrame(
        {
            "event": event,
            "timestamp": active.index,
            "network_mean_mm": active.mean(axis=1, skipna=True).values,
            "network_median_mm": active.median(axis=1, skipna=True).values,
            "network_max_mm": active.max(axis=1, skipna=True).values,
            "reporting_sensors": active.count(axis=1).values,
            "active_sensors": len(active_stations),
        }
    )
    hourly["date"] = hourly["timestamp"].dt.floor("D")
    hourly["hour"] = hourly["timestamp"].dt.hour

    daily_sensor_rows: list[dict[str, Any]] = []
    dates = pd.date_range(active.index.min().floor("D"), active.index.max().floor("D"), freq="D")
    minimum_valid_hours = math.ceil(24 * settings.minimum_daily_completeness)

    for date in dates:
        day_slice = active.loc[date : date + pd.Timedelta(hours=23)]
        for station in active_stations:
            valid_hours = int(day_slice[station].count())
            qualifies = valid_hours >= minimum_valid_hours
            daily_total = (
                float(day_slice[station].sum()) if qualifies else np.nan
            )
            daily_sensor_rows.append(
                {
                    "event": event,
                    "day_number": int((date - dates[0]).days + 1),
                    "date": date,
                    "station": station,
                    "daily_accumulation_mm": daily_total,
                    "valid_hours": valid_hours,
                    "expected_hours": 24,
                    "qualifies_for_daily_average": qualifies,
                }
            )

    daily_sensor = pd.DataFrame(daily_sensor_rows)
    daily_rows: list[dict[str, Any]] = []
    for (date, day_number), group in daily_sensor.groupby(
        ["date", "day_number"], sort=True
    ):
        qualifying = group.loc[
            group["qualifies_for_daily_average"], "daily_accumulation_mm"
        ].dropna()
        hour_group = hourly.loc[hourly["date"] == date]
        daily_rows.append(
            {
                "event": event,
                "day_number": day_number,
                "date": date,
                "network_average_daily_accumulation_mm": qualifying.mean(),
                "network_median_daily_accumulation_mm": qualifying.median(),
                "minimum_station_daily_accumulation_mm": qualifying.min(),
                "maximum_station_daily_accumulation_mm": qualifying.max(),
                "station_daily_standard_deviation_mm": qualifying.std(ddof=1),
                "qualifying_sensors": int(qualifying.count()),
                "active_sensors": len(active_stations),
                "sum_of_hourly_network_means_mm": hour_group["network_mean_mm"].sum(
                    min_count=minimum_valid_hours
                ),
            }
        )
    daily = pd.DataFrame(daily_rows)

    valid_hourly = hourly.dropna(subset=["network_mean_mm"])
    peak_hour_row = (
        valid_hourly.loc[valid_hourly["network_mean_mm"].idxmax()]
        if not valid_hourly.empty
        else None
    )

    station_maxima = active.max(axis=0, skipna=True)
    if station_maxima.notna().any():
        peak_station_name = str(station_maxima.idxmax())
        peak_station_timestamp = active[peak_station_name].idxmax()
        peak_station_value = float(active.at[peak_station_timestamp, peak_station_name])
    else:
        peak_station_timestamp = pd.NaT
        peak_station_name = ""
        peak_station_value = np.nan

    wettest_row = (
        daily.loc[daily["network_average_daily_accumulation_mm"].idxmax()]
        if daily["network_average_daily_accumulation_mm"].notna().any()
        else None
    )
    possible_cells = len(active.index) * len(active_stations)
    valid_cells = int(active.count().sum())

    summary = {
        "event": event,
        "start_timestamp": active.index.min(),
        "end_timestamp": active.index.max(),
        "calendar_days": int(len(daily)),
        "input_sensors": int(len(sensor_status)),
        "active_sensors": int(len(active_stations)),
        "excluded_sensors": int((~sensor_status["included_in_analysis"]).sum()),
        "omitted_cell_values": int(len(audit)),
        "statistical_high_outliers": int(
            sum(
                row["reason"] == "statistical_high_outlier_spatial_and_temporal"
                for row in audit
            )
        ),
        "valid_cell_completeness_percent": (
            100.0 * valid_cells / possible_cells if possible_cells else np.nan
        ),
        "event_total_network_average_mm": daily[
            "network_average_daily_accumulation_mm"
        ].sum(min_count=len(daily)),
        "mean_daily_network_average_mm": daily[
            "network_average_daily_accumulation_mm"
        ].mean(),
        "median_daily_network_average_mm": daily[
            "network_average_daily_accumulation_mm"
        ].median(),
        "wettest_day": wettest_row["date"] if wettest_row is not None else pd.NaT,
        "wettest_day_network_average_mm": (
            wettest_row["network_average_daily_accumulation_mm"]
            if wettest_row is not None
            else np.nan
        ),
        "peak_network_mean_hour_timestamp": (
            peak_hour_row["timestamp"] if peak_hour_row is not None else pd.NaT
        ),
        "peak_network_mean_hour_mm": (
            peak_hour_row["network_mean_mm"] if peak_hour_row is not None else np.nan
        ),
        "highest_clean_station_hour_timestamp": peak_station_timestamp,
        "highest_clean_station_hour_station": peak_station_name,
        "highest_clean_station_hour_mm": peak_station_value,
    }
    return hourly, daily_sensor, daily, summary


# ---------------------------------------------------------------------------
# PLOTS
# ---------------------------------------------------------------------------


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
            "axes.grid": False,
        }
    )


def add_hourly_intensity_lines(ax: plt.Axes, data_max: float) -> None:
    top = max(3.0, data_max * 1.15 if np.isfinite(data_max) else 3.0)
    for threshold, label, color in HOURLY_INTENSITY_LINES:
        if threshold <= top:
            ax.axhline(
                threshold,
                color=color,
                linestyle="--",
                linewidth=1.0,
                alpha=0.85,
                label=f"{label} ({threshold:g} mm/h)",
            )


def finish_figure(fig: plt.Figure, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def plot_hourly_calendar_days(hourly: pd.DataFrame, plots_root: Path) -> None:
    output_dir = plots_root / "hourly_by_calendar_day"
    for (event, date), group in hourly.groupby(["event", "date"], sort=True):
        group = group.sort_values("hour")
        values = group.set_index("hour")["network_mean_mm"].reindex(range(24))
        color = EVENT_COLORS.get(event, "#2878B5")

        fig, ax = plt.subplots(figsize=(10.5, 5.2))
        ax.bar(range(24), values.values, color=color, width=0.78, edgecolor="none")
        data_max = float(values.max()) if values.notna().any() else 0.0
        add_hourly_intensity_lines(ax, data_max)
        ax.set_title(
            f"{event}: Hourly Network-Mean Rainfall — {pd.Timestamp(date):%d %b %Y}"
        )
        ax.set_xlabel("Hour (PHT)")
        ax.set_ylabel("Network-mean rainfall (mm/h)")
        ax.set_xticks(range(0, 24, 2))
        ax.set_xticklabels([f"{hour:02d}:00" for hour in range(0, 24, 2)])
        ax.set_xlim(-0.7, 23.7)
        ax.set_ylim(bottom=0)
        handles, labels = ax.get_legend_handles_labels()
        if handles:
            ax.legend(
                frameon=True,
                facecolor="white",
                edgecolor="none",
                framealpha=0.90,
                fontsize=8,
                ncol=2,
                loc="upper right",
            )

        file_name = f"{event}_{pd.Timestamp(date):%Y-%m-%d}_hourly.png"
        finish_figure(fig, output_dir / file_name)


def annotate_bars(ax: plt.Axes, bars: Any, decimals: int = 1) -> None:
    for bar in bars:
        height = bar.get_height()
        if not np.isfinite(height):
            continue
        ax.annotate(
            f"{height:.{decimals}f}",
            (bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 4),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=9,
        )


def plot_daily_by_event(daily: pd.DataFrame, plots_root: Path) -> None:
    output_dir = plots_root / "daily_by_event"
    for event, group in daily.groupby("event", sort=False):
        group = group.sort_values("day_number")
        x = np.arange(len(group))
        values = group["network_average_daily_accumulation_mm"].to_numpy(float)
        labels = [f"Day {number}\n{date:%d %b}" for number, date in zip(
            group["day_number"], pd.to_datetime(group["date"])
        )]

        fig, ax = plt.subplots(figsize=(7.5, 5.0))
        bars = ax.bar(
            x,
            values,
            color=EVENT_COLORS.get(event, "#2878B5"),
            width=0.62,
        )
        annotate_bars(ax, bars)
        ax.set_title(f"{event}: Daily Rainfall Averaged Across Sensors")
        ax.set_xlabel("Event day")
        ax.set_ylabel("Network-average daily accumulation (mm)")
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        ax.set_ylim(bottom=0)
        finish_figure(fig, output_dir / f"{event}_daily_accumulation.png")


def plot_successive_day_hourly_comparisons(
    hourly: pd.DataFrame,
    plots_root: Path,
) -> None:
    output_dir = plots_root / "successive_day_hourly_comparisons"
    events = list(hourly["event"].drop_duplicates())
    day_map = (
        hourly[["event", "date"]]
        .drop_duplicates()
        .sort_values(["event", "date"])
        .assign(day_number=lambda frame: frame.groupby("event").cumcount() + 1)
    )
    with_days = hourly.merge(day_map, on=["event", "date"], how="left")

    for day_number in sorted(with_days["day_number"].dropna().unique()):
        comparison = with_days.loc[with_days["day_number"] == day_number]
        x = np.arange(24)
        width = 0.38 if len(events) == 2 else 0.8 / max(len(events), 1)

        fig, ax = plt.subplots(figsize=(11.5, 5.5))
        all_values: list[float] = []
        for event_index, event in enumerate(events):
            event_data = comparison.loc[comparison["event"] == event].sort_values("hour")
            values = event_data.set_index("hour")["network_mean_mm"].reindex(range(24))
            all_values.extend(values.dropna().tolist())
            date_text = (
                f"{pd.Timestamp(event_data['date'].iloc[0]):%d %b %Y}"
                if not event_data.empty
                else "no date"
            )
            offset = (event_index - (len(events) - 1) / 2) * width
            ax.bar(
                x + offset,
                values.values,
                width=width,
                color=EVENT_COLORS.get(event),
                label=f"{event} — {date_text}",
                alpha=0.92,
            )

        add_hourly_intensity_lines(ax, max(all_values) if all_values else 0.0)
        ax.set_title(f"Successive-Day Hourly Comparison — Day {int(day_number)}")
        ax.set_xlabel("Hour (PHT)")
        ax.set_ylabel("Network-mean rainfall (mm/h)")
        ax.set_xticks(range(0, 24, 2))
        ax.set_xticklabels([f"{hour:02d}:00" for hour in range(0, 24, 2)])
        ax.set_xlim(-0.8, 23.8)
        ax.set_ylim(bottom=0)
        ax.legend(
            frameon=True,
            facecolor="white",
            edgecolor="none",
            framealpha=0.90,
            fontsize=8,
            ncol=2,
            loc="upper right",
        )
        finish_figure(
            fig,
            output_dir / f"day_{int(day_number):02d}_hourly_comparison.png",
        )


def plot_successive_day_daily_comparison(
    daily: pd.DataFrame,
    plots_root: Path,
) -> None:
    output_dir = plots_root / "successive_day_daily_comparison"
    events = list(daily["event"].drop_duplicates())
    day_numbers = sorted(daily["day_number"].unique())
    x = np.arange(len(day_numbers))
    width = 0.38 if len(events) == 2 else 0.8 / max(len(events), 1)

    fig, ax = plt.subplots(figsize=(9.0, 5.4))
    for event_index, event in enumerate(events):
        group = daily.loc[daily["event"] == event].set_index("day_number")
        values = group["network_average_daily_accumulation_mm"].reindex(day_numbers)
        date_labels = {
            day: pd.Timestamp(group.loc[day, "date"]).strftime("%d %b")
            for day in group.index
        }
        legend_dates = ", ".join(
            f"D{day} {date_labels[day]}" for day in day_numbers if day in date_labels
        )
        offset = (event_index - (len(events) - 1) / 2) * width
        bars = ax.bar(
            x + offset,
            values.values,
            width=width,
            color=EVENT_COLORS.get(event),
            label=f"{event} ({legend_dates})",
        )
        annotate_bars(ax, bars)

    ax.set_title("Successive-Day Daily Rainfall Comparison")
    ax.set_xlabel("Relative event day")
    ax.set_ylabel("Network-average daily accumulation (mm)")
    ax.set_xticks(x)
    ax.set_xticklabels([f"Day {int(day)}" for day in day_numbers])
    ax.set_ylim(bottom=0)
    ax.legend(frameon=False, fontsize=8)
    finish_figure(fig, output_dir / "successive_day_daily_comparison.png")


# ---------------------------------------------------------------------------
# OUTPUT TABLES AND SUMMARY
# ---------------------------------------------------------------------------


def fmt_mm(value: Any) -> str:
    return "not available" if pd.isna(value) else f"{float(value):.2f} mm"


def fmt_timestamp(value: Any) -> str:
    return "not available" if pd.isna(value) else pd.Timestamp(value).strftime("%d %b %Y %H:%M PHT")


def percent_difference(new_value: float, reference_value: float) -> str:
    if pd.isna(new_value) or pd.isna(reference_value) or reference_value == 0:
        return "percentage difference unavailable"
    percent = 100.0 * (new_value - reference_value) / reference_value
    direction = "higher" if percent > 0 else "lower" if percent < 0 else "equal"
    return f"{abs(percent):.1f}% {direction}"


def write_plain_language_summary(
    output_path: Path,
    event_summaries: pd.DataFrame,
    daily: pd.DataFrame,
    sensor_status: pd.DataFrame,
    settings: OutlierSettings,
) -> None:
    lines = [
        "CDEHab vs MDHab Rainfall Event Comparison",
        "===========================================",
        "",
        "Method",
        "------",
        (
            "Hourly statistics are means across active reporting sensors. Daily rainfall is "
            "calculated by accumulating each qualifying sensor-day first and then averaging "
            "those daily totals across sensors. A sensor-day must retain at least "
            f"{settings.minimum_daily_completeness:.0%} of its 24 expected hourly values. "
            "Flagged readings are omitted and are not imputed."
        ),
        (
            "Invalid/physically inadmissible readings are removed automatically. Statistical "
            "high outliers are removed only when they are unusually high both relative to the "
            "network at that hour and relative to the same sensor's surrounding hours. Zero "
            "rainfall remains valid, except that an all-zero sensor over the entire event is "
            "treated as inactive by default. Review outlier_audit.csv and sensor_status.csv "
            "before using the results operationally or in a publication."
        ),
        (
            "The two events use different sensor rosters. The results therefore compare "
            "network-wide averages, not matched readings from an identical set of stations."
        ),
        "",
        "Event statistics",
        "----------------",
    ]

    for _, row in event_summaries.iterrows():
        event = row["event"]
        excluded_statuses = sensor_status.loc[
            (sensor_status["event"] == event)
            & (~sensor_status["included_in_analysis"]),
            "status",
        ].value_counts()
        status_text = (
            ", ".join(f"{status}: {count}" for status, count in excluded_statuses.items())
            if not excluded_statuses.empty
            else "none"
        )
        lines.extend(
            [
                f"{event}",
                f"  Period: {fmt_timestamp(row['start_timestamp'])} to {fmt_timestamp(row['end_timestamp'])}",
                f"  Sensors: {int(row['active_sensors'])} active of {int(row['input_sensors'])}; excluded — {status_text}",
                f"  Clean-data completeness among active sensors: {row['valid_cell_completeness_percent']:.1f}%",
                f"  Event-total network-average rainfall: {fmt_mm(row['event_total_network_average_mm'])}",
                f"  Mean daily network-average rainfall: {fmt_mm(row['mean_daily_network_average_mm'])}",
                f"  Median daily network-average rainfall: {fmt_mm(row['median_daily_network_average_mm'])}",
                f"  Wettest day: {pd.Timestamp(row['wettest_day']):%d %b %Y} ({fmt_mm(row['wettest_day_network_average_mm'])})",
                f"  Peak network-mean hour: {fmt_timestamp(row['peak_network_mean_hour_timestamp'])} ({fmt_mm(row['peak_network_mean_hour_mm'])})",
                (
                    "  Highest retained station-hour: "
                    f"{row['highest_clean_station_hour_station']}, "
                    f"{fmt_timestamp(row['highest_clean_station_hour_timestamp'])} "
                    f"({fmt_mm(row['highest_clean_station_hour_mm'])})"
                ),
                (
                    f"  Omitted cell values: {int(row['omitted_cell_values'])}, including "
                    f"{int(row['statistical_high_outliers'])} statistical high outlier(s)"
                ),
                "",
            ]
        )

    if len(event_summaries) >= 2:
        reference = event_summaries.iloc[0]
        comparison = event_summaries.iloc[1]
        event_total_difference = (
            comparison["event_total_network_average_mm"]
            - reference["event_total_network_average_mm"]
        )
        mean_daily_difference = (
            comparison["mean_daily_network_average_mm"]
            - reference["mean_daily_network_average_mm"]
        )
        peak_difference = (
            comparison["peak_network_mean_hour_mm"]
            - reference["peak_network_mean_hour_mm"]
        )
        lines.extend(
            [
                "Direct comparison",
                "-----------------",
                (
                    f"{comparison['event']} recorded an event-total network-average rainfall "
                    f"of {fmt_mm(comparison['event_total_network_average_mm'])}, compared with "
                    f"{fmt_mm(reference['event_total_network_average_mm'])} for {reference['event']}. "
                    f"The difference was {event_total_difference:+.2f} mm "
                    f"({percent_difference(comparison['event_total_network_average_mm'], reference['event_total_network_average_mm'])})."
                ),
                (
                    f"Mean daily network-average rainfall differed by {mean_daily_difference:+.2f} mm/day; "
                    f"the peak network-mean hour differed by {peak_difference:+.2f} mm/h."
                ),
                "",
                "Successive-day daily values",
                "---------------------------",
            ]
        )

        pivot = daily.pivot(
            index="day_number",
            columns="event",
            values="network_average_daily_accumulation_mm",
        )
        for day_number, values in pivot.iterrows():
            parts = [
                f"{event}: {fmt_mm(values.get(event, np.nan))}"
                for event in event_summaries["event"]
            ]
            lines.append(f"Day {int(day_number)} — " + "; ".join(parts))

    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def prepare_csv_frame(frame: pd.DataFrame) -> pd.DataFrame:
    output = frame.copy()
    for column in output.columns:
        if pd.api.types.is_datetime64_any_dtype(output[column]):
            if "timestamp" in column:
                output[column] = output[column].dt.strftime("%Y-%m-%d %H:%M")
            else:
                output[column] = output[column].dt.strftime("%Y-%m-%d")
    return output


def save_csv(frame: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    prepare_csv_frame(frame).to_csv(
        output_path,
        index=False,
        encoding="utf-8-sig",
        float_format="%.3f",
    )


def write_all_outputs(
    output_directory: Path,
    cleaned_by_event: dict[str, pd.DataFrame],
    hourly: pd.DataFrame,
    daily_sensor: pd.DataFrame,
    daily: pd.DataFrame,
    event_summaries: pd.DataFrame,
    outlier_audit: pd.DataFrame,
    sensor_status: pd.DataFrame,
    data_issues: pd.DataFrame,
    settings: OutlierSettings,
) -> None:
    output_directory.mkdir(parents=True, exist_ok=True)
    tables_dir = output_directory / "Tables"
    plots_dir = output_directory / "Plots"

    for event, cleaned in cleaned_by_event.items():
        cleaned_output = cleaned.reset_index()
        cleaned_output["Date & Time"] = cleaned_output["Date & Time"].dt.strftime(
            "%Y-%m-%d %H:%M"
        )
        save_csv(
            cleaned_output,
            tables_dir / f"cleaned_hourly_{event}.csv",
        )

    save_csv(hourly, tables_dir / "hourly_network_statistics.csv")
    save_csv(daily_sensor, tables_dir / "daily_station_accumulations.csv")
    save_csv(daily, tables_dir / "daily_network_statistics.csv")
    save_csv(event_summaries, tables_dir / "event_summary_statistics.csv")
    save_csv(outlier_audit, tables_dir / "outlier_audit.csv")
    save_csv(sensor_status, tables_dir / "sensor_status.csv")
    save_csv(data_issues, tables_dir / "data_issues.csv")

    apply_plot_style()
    plot_hourly_calendar_days(hourly, plots_dir)
    plot_daily_by_event(daily, plots_dir)
    plot_successive_day_hourly_comparisons(hourly, plots_dir)
    plot_successive_day_daily_comparison(daily, plots_dir)

    write_plain_language_summary(
        output_directory / "CDEHab_vs_MDHab_summary.txt",
        event_summaries,
        daily,
        sensor_status,
        settings,
    )


# ---------------------------------------------------------------------------
# PROGRAM ENTRY POINT
# ---------------------------------------------------------------------------


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Compare the CDEHab and MDHab hourly rainfall events."
    )
    parser.add_argument("--cdehab", type=Path, default=CDEHAB_FILE)
    parser.add_argument("--mdhab", type=Path, default=MDHAB_FILE)
    parser.add_argument("--output", type=Path, default=OUTPUT_DIRECTORY)
    parser.add_argument(
        "--disable-statistical-outliers",
        action="store_true",
        help="Only remove invalid values and values above the hard limit.",
    )
    parser.add_argument(
        "--keep-all-zero-sensors",
        action="store_true",
        help="Retain sensors that report zero for the entire event.",
    )
    return parser


def run_analysis(
    event_files: list[tuple[str, Path]],
    output_directory: Path,
    settings: OutlierSettings,
) -> None:
    all_hourly: list[pd.DataFrame] = []
    all_daily_sensor: list[pd.DataFrame] = []
    all_daily: list[pd.DataFrame] = []
    all_summaries: list[dict[str, Any]] = []
    all_sensor_status: list[pd.DataFrame] = []
    all_outlier_audit: list[dict[str, Any]] = []
    all_data_issues: list[dict[str, str]] = []
    cleaned_by_event: dict[str, pd.DataFrame] = {}

    for event, file_path in event_files:
        print(f"Reading {event}: {file_path}")
        rainfall, initial_audit, data_issues = load_event(
            event, file_path, settings
        )
        sensor_status = identify_sensor_status(event, rainfall, settings)
        active_stations = sensor_status.loc[
            sensor_status["included_in_analysis"], "station"
        ].tolist()
        cleaned, statistical_audit = omit_statistical_outliers(
            event, rainfall, active_stations, settings
        )
        event_audit = initial_audit + statistical_audit
        hourly, daily_sensor, daily, summary = calculate_event_statistics(
            event,
            cleaned,
            sensor_status,
            event_audit,
            settings,
        )

        cleaned_by_event[event] = cleaned
        all_hourly.append(hourly)
        all_daily_sensor.append(daily_sensor)
        all_daily.append(daily)
        all_summaries.append(summary)
        all_sensor_status.append(sensor_status)
        all_outlier_audit.extend(event_audit)
        all_data_issues.extend(data_issues)

    hourly = pd.concat(all_hourly, ignore_index=True)
    daily_sensor = pd.concat(all_daily_sensor, ignore_index=True)
    daily = pd.concat(all_daily, ignore_index=True)
    event_summaries = pd.DataFrame(all_summaries)
    sensor_status = pd.concat(all_sensor_status, ignore_index=True)
    outlier_audit = pd.DataFrame(all_outlier_audit, columns=OUTLIER_AUDIT_COLUMNS)
    data_issues = pd.DataFrame(all_data_issues, columns=DATA_ISSUE_COLUMNS)

    write_all_outputs(
        output_directory=output_directory,
        cleaned_by_event=cleaned_by_event,
        hourly=hourly,
        daily_sensor=daily_sensor,
        daily=daily,
        event_summaries=event_summaries,
        outlier_audit=outlier_audit,
        sensor_status=sensor_status,
        data_issues=data_issues,
        settings=settings,
    )

    print(f"\nAnalysis complete. Outputs written to:\n{output_directory}")


def main() -> int:
    args = build_argument_parser().parse_args()
    settings = replace(
        DEFAULT_SETTINGS,
        statistical_screening=not args.disable_statistical_outliers,
        exclude_all_zero_sensors=not args.keep_all_zero_sensors,
    )

    try:
        run_analysis(
            [("CDEHab", args.cdehab), ("MDHab", args.mdhab)],
            args.output,
            settings,
        )
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
