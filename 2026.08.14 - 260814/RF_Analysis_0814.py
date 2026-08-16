from __future__ import annotations

"""
RF_Analysis_0814.py

Rainfall analysis for 14 August 2026, 00:00-12:00.

Outputs
-------
1) Bar plot of hourly network-mean rainfall with intensity thresholds.
2) Plain-text rainfall report.
3) Peak-hour rainfall intensity map using rain-gauge points plus AWS points
   extracted from the Early Warning Systems shapefile.
4) CSV of GIS-to-Excel station matching diagnostics for the peak-hour map.

Notes
-----
- Duplicate Excel station columns such as "Station" and "Station.1" are
  consolidated by taking their row-wise mean before network statistics are
  calculated. This prevents duplicated stations from being overweighted.
- The "network peak hour" is the timestamp with the highest ordinary mean
  rainfall across unique stations. The map uses this hour.
- The report separately identifies the highest individual hourly station
  observation during the period.
- The analysis window is inclusive of both 00:00 and 12:00 timestamps.
"""

from pathlib import Path
import difflib
import math
import re
import sys
import warnings

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import BoundaryNorm, ListedColormap
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

try:
    import geopandas as gpd
except ImportError as exc:
    raise SystemExit(
        "geopandas is required for the map. Install it with: pip install geopandas"
    ) from exc

try:
    from shapely import contains_xy
except ImportError:  # Shapely < 2 fallback is handled later.
    contains_xy = None

try:
    from shapely.geometry import Point
except ImportError as exc:
    raise SystemExit("shapely is required for the map.") from exc


# =============================================================================
# CONFIGURATION
# =============================================================================

INPUT_XLSX = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.14 - 260814\Processed Data\RF_260814.xlsx"
)
OUTPUT_DIR = Path(
    r"C:\Users\QCUSER\Documents\Analysis\2026.08.14 - 260814\Outputs"
)

WATERWAYS_SHP = Path(
    r"C:\Users\QCUSER\Documents\Given\QC Waterways\Waterways_2020_UTM.shp"
)
RAIN_GAUGES_SHP = Path(
    r"C:\Users\QCUSER\Documents\Given\RainGauges_2026\Rain Gauges (July 2026).shp"
)
EWS_SHP = Path(
    r"C:\Users\QCUSER\Documents\Given\EWS_2025\EarlyWarningSystems_2025.shp"
)
BARANGAYS_SHP = Path(
    r"C:\Users\QCUSER\Documents\Given\BrgyBoundary\BarangayBoundary.shp"
)

ANALYSIS_DATE = pd.Timestamp("2026-08-14")
START_TIME = "00:00"
END_TIME = "12:00"
END_INCLUSIVE = True

# Match tolerance for GIS station labels -> Excel station headers.
FUZZY_MATCH_CUTOFF = 0.72

# IDW surface resolution. 220 x 220 is smooth enough for a citywide quick-look map.
IDW_GRID_SIZE = 220
IDW_POWER = 2.0

# Plot thresholds (mm/hr).
THRESHOLDS = [
    (2.5, "Moderate", "orange"),
    (7.5, "Heavy", "red"),
    (15.0, "Intense", "darkred"),
    (30.0, "Torrential", "violet"),
]


# =============================================================================
# GENERAL HELPERS
# =============================================================================


def ensure_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} not found:\n  {path}")


def safe_read_vector(path: Path, label: str) -> gpd.GeoDataFrame:
    ensure_exists(path, label)
    gdf = gpd.read_file(path)
    if gdf.empty:
        raise ValueError(f"{label} is empty: {path}")
    if gdf.crs is None:
        raise ValueError(f"{label} has no CRS defined: {path}")

    # Repair common invalid geometry problems without changing valid features.
    bad = ~gdf.geometry.is_valid & gdf.geometry.notna()
    if bad.any():
        warnings.warn(f"{label}: repairing {int(bad.sum())} invalid geometries with buffer(0).")
        gdf.loc[bad, "geometry"] = gdf.loc[bad, "geometry"].buffer(0)

    return gdf[gdf.geometry.notna() & ~gdf.geometry.is_empty].copy()


def parse_numeric(series: pd.Series) -> pd.Series:
    """Convert Excel values to float, including decimal-comma strings such as '2,5'."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    s = series.astype("string").str.strip()
    s = s.replace({"": pd.NA, "-": pd.NA, "--": pd.NA, "N/A": pd.NA, "NA": pd.NA})

    # Handle decimal commas while retaining ordinary decimal points.
    has_comma = s.str.contains(",", regex=False, na=False)
    s.loc[has_comma] = s.loc[has_comma].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)

    return pd.to_numeric(s, errors="coerce")


def strip_pandas_duplicate_suffix(name: str) -> str:
    """Pandas renames duplicate Excel headers as Name.1, Name.2, etc."""
    return re.sub(r"\.\d+$", "", str(name).strip())


def normalize_station_name(value: object) -> str:
    """Normalize station/site names for robust Excel <-> shapefile matching."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return ""

    s = str(value).lower().strip()
    s = s.replace("ñ", "n")
    s = re.sub(r"[’'`]+", "", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)

    # Normalize common abbreviations found in QC station names.
    replacements = {
        r"\bbrgy\b": "barangay",
        r"\bbrg\b": "barangay",
        r"\bbrgyhall\b": "barangay hall",
        r"\bes\b": "elementary school",
        r"\bhs\b": "high school",
        r"\bsr\b": "senior",
        r"\bqcu\b": "quezon city university",
        r"\bqc\b": "quezon city",
        r"\bmp\b": "multipurpose",
    }
    for pattern, repl in replacements.items():
        s = re.sub(pattern, repl, s)

    # Common spelling standardization.
    s = s.replace("multi purpose", "multipurpose")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def station_similarity(a: str, b: str) -> float:
    """Blend sequence similarity with token overlap."""
    a_n = normalize_station_name(a)
    b_n = normalize_station_name(b)
    if not a_n or not b_n:
        return 0.0
    if a_n == b_n:
        return 1.0

    seq = difflib.SequenceMatcher(None, a_n, b_n).ratio()
    ta, tb = set(a_n.split()), set(b_n.split())
    jaccard = len(ta & tb) / len(ta | tb) if (ta | tb) else 0.0

    # Containment is useful for names with qualifiers such as "Brgy Hall".
    containment = 0.0
    if a_n in b_n or b_n in a_n:
        containment = min(len(a_n), len(b_n)) / max(len(a_n), len(b_n))

    return max(seq, 0.65 * seq + 0.35 * jaccard, containment)


def match_to_excel_name(gis_name: object, excel_names: list[str]) -> tuple[str | None, float, str]:
    """Return (matched Excel station, score, method)."""
    gis_norm = normalize_station_name(gis_name)
    if not gis_norm:
        return None, 0.0, "blank"

    exact = {normalize_station_name(x): x for x in excel_names}
    if gis_norm in exact:
        return exact[gis_norm], 1.0, "exact"

    scored = [(station_similarity(str(gis_name), x), x) for x in excel_names]
    score, best = max(scored, default=(0.0, None))
    if best is not None and score >= FUZZY_MATCH_CUTOFF:
        return best, float(score), "fuzzy"

    return None, float(score), "unmatched"


# =============================================================================
# RAINFALL DATA
# =============================================================================


def load_rainfall() -> tuple[pd.DataFrame, dict[str, list[str]]]:
    ensure_exists(INPUT_XLSX, "Rainfall workbook")

    raw = pd.read_excel(INPUT_XLSX, sheet_name=0)
    if raw.shape[1] < 2:
        raise ValueError("Rainfall workbook must contain a Date & Time column and at least one station column.")

    date_col = raw.columns[0]
    timestamps = pd.to_datetime(raw[date_col], errors="coerce")
    if timestamps.isna().all():
        raise ValueError(f"Could not parse timestamps from first column: {date_col!r}")

    values = raw.drop(columns=[date_col]).copy()
    values = values.apply(parse_numeric)

    # Collapse duplicate station columns created by Excel/pandas, e.g. Station + Station.1.
    groups: dict[str, list[str]] = {}
    for col in values.columns:
        base = strip_pandas_duplicate_suffix(col)
        groups.setdefault(base, []).append(col)

    collapsed = pd.DataFrame(index=values.index)
    for base, cols in groups.items():
        if len(cols) == 1:
            collapsed[base] = values[cols[0]]
        else:
            collapsed[base] = values[cols].mean(axis=1, skipna=True)

    collapsed.insert(0, "Date & Time", timestamps)
    collapsed = collapsed.dropna(subset=["Date & Time"]).sort_values("Date & Time").reset_index(drop=True)

    duplicate_groups = {k: v for k, v in groups.items() if len(v) > 1}
    return collapsed, duplicate_groups


def select_period(df: pd.DataFrame) -> pd.DataFrame:
    start = ANALYSIS_DATE + pd.Timedelta(START_TIME + ":00")
    end = ANALYSIS_DATE + pd.Timedelta(END_TIME + ":00")

    if END_INCLUSIVE:
        mask = (df["Date & Time"] >= start) & (df["Date & Time"] <= end)
    else:
        mask = (df["Date & Time"] >= start) & (df["Date & Time"] < end)

    period = df.loc[mask].copy()
    if period.empty:
        raise ValueError(f"No rainfall rows found from {start} to {end}.")
    return period


def calculate_statistics(period: pd.DataFrame) -> dict:
    station_cols = [c for c in period.columns if c != "Date & Time"]
    rain = period.set_index("Date & Time")[station_cols]

    # Ordinary network mean for each timestamp, using all valid unique stations.
    hourly_mean = rain.mean(axis=1, skipna=True)
    hourly_n = rain.notna().sum(axis=1)

    # Mean accumulated rainfall: accumulate each unique station over the period,
    # then take the ordinary mean of station accumulations.
    station_accum = rain.sum(axis=0, skipna=True, min_count=1)
    station_obs_n = rain.notna().sum(axis=0)
    mean_accum = float(station_accum.mean(skipna=True))

    peak_network_time = hourly_mean.idxmax()
    peak_network_mean = float(hourly_mean.loc[peak_network_time])

    stacked = rain.stack(dropna=True)
    if stacked.empty:
        raise ValueError("The selected period contains no numeric rainfall observations.")
    peak_pair = stacked.idxmax()  # (timestamp, station)
    peak_station_time = pd.Timestamp(peak_pair[0])
    peak_station = str(peak_pair[1])
    peak_station_value = float(stacked.loc[peak_pair])

    max_accum_station = station_accum.idxmax()
    max_accum_value = float(station_accum.loc[max_accum_station])

    peak_hour_values = rain.loc[peak_network_time].sort_values(ascending=False)

    return {
        "rain": rain,
        "hourly_mean": hourly_mean,
        "hourly_n": hourly_n,
        "station_accum": station_accum,
        "station_obs_n": station_obs_n,
        "mean_accum": mean_accum,
        "peak_network_time": peak_network_time,
        "peak_network_mean": peak_network_mean,
        "peak_station_time": peak_station_time,
        "peak_station": peak_station,
        "peak_station_value": peak_station_value,
        "max_accum_station": str(max_accum_station),
        "max_accum_value": max_accum_value,
        "peak_hour_values": peak_hour_values,
    }


# =============================================================================
# HOURLY BAR PLOT
# =============================================================================


def plot_hourly_mean(stats: dict) -> Path:
    hourly = stats["hourly_mean"]
    out = OUTPUT_DIR / "RF_260814_hourly_mean_0000-1200.png"

    fig, ax = plt.subplots(figsize=(13, 7.2))
    labels = [t.strftime("%I %p").lstrip("0") for t in hourly.index]
    x = np.arange(len(hourly))

    bars = ax.bar(x, hourly.values, edgecolor="black", linewidth=0.6)

    for value, label, color in THRESHOLDS:
        ax.axhline(value, linestyle="--", linewidth=1.7, color=color, label=f"{label} ({value:g} mm/hr)")

    # Give room for threshold lines and value labels without flattening the bars.
    ymax = max(32.0, float(np.nanmax(hourly.values)) * 1.18)
    ax.set_ylim(0, ymax)

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right")
    ax.set_ylabel("Mean hourly rainfall (mm)")
    ax.set_xlabel("14 August 2026 (PHT)")
    ax.set_title("Quezon City Mean Hourly Rainfall | 00:00-12:00, 14 August 2026")
    ax.grid(axis="y", alpha=0.25)

    for bar, value in zip(bars, hourly.values):
        if np.isfinite(value):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + ymax * 0.012,
                f"{value:.2f}",
                ha="center",
                va="bottom",
                fontsize=8,
                rotation=90 if len(hourly) > 10 else 0,
            )

    ax.legend(loc="upper left", frameon=True)
    fig.tight_layout()
    fig.savefig(out, dpi=220, bbox_inches="tight")
    plt.close(fig)
    return out


# =============================================================================
# GIS / AWS HELPERS
# =============================================================================


def filter_aws(ews: gpd.GeoDataFrame) -> tuple[gpd.GeoDataFrame, list[str]]:
    """
    Identify Automated Weather Station records by searching all text fields.
    Returns the filtered AWS rows and text fields searched.
    """
    text_cols = [
        c for c in ews.columns
        if c != ews.geometry.name and (
            pd.api.types.is_object_dtype(ews[c]) or pd.api.types.is_string_dtype(ews[c])
        )
    ]

    if not text_cols:
        warnings.warn("EWS shapefile has no text fields; no AWS features can be identified automatically.")
        return ews.iloc[0:0].copy(), []

    pattern = re.compile(
        r"(?:\bAWS\b|AUTOMATED\s+WEATHER\s+STATION|AUTOMATIC\s+WEATHER\s+STATION|WEATHER\s+STATION)",
        flags=re.IGNORECASE,
    )

    mask = pd.Series(False, index=ews.index)
    for col in text_cols:
        mask |= ews[col].astype("string").str.contains(pattern, na=False)

    aws = ews.loc[mask].copy()
    if aws.empty:
        warnings.warn(
            "No AWS records were found in the EWS shapefile using the keywords AWS / Automated Weather Station / Weather Station."
        )
    return aws, text_cols


def choose_station_name_field(gdf: gpd.GeoDataFrame, excel_names: list[str]) -> tuple[str | None, dict[str, int]]:
    """
    Pick the most plausible station-name field by measuring how many values can
    be matched to Excel station headers.
    """
    candidate_cols = [
        c for c in gdf.columns
        if c != gdf.geometry.name and (
            pd.api.types.is_object_dtype(gdf[c]) or pd.api.types.is_string_dtype(gdf[c])
        )
    ]
    if not candidate_cols:
        return None, {}

    scores: dict[str, int] = {}
    for col in candidate_cols:
        count = 0
        vals = gdf[col].dropna().astype(str).drop_duplicates().head(500)
        for value in vals:
            matched, score, _ = match_to_excel_name(value, excel_names)
            if matched is not None and score >= FUZZY_MATCH_CUTOFF:
                count += 1
        scores[col] = count

    # Prefer semantically likely fields if match scores tie.
    semantic_rank = {}
    for col in candidate_cols:
        c = col.lower()
        semantic_rank[col] = sum(k in c for k in ["station", "name", "site", "location", "facility"])

    best = max(candidate_cols, key=lambda c: (scores.get(c, 0), semantic_rank.get(c, 0)))
    return best, scores


def prepare_station_layer(
    gdf: gpd.GeoDataFrame,
    source: str,
    excel_names: list[str],
) -> tuple[gpd.GeoDataFrame, str | None, dict[str, int]]:
    if gdf.empty:
        empty = gdf.copy()
        empty["source"] = source
        empty["gis_station_name"] = pd.Series(dtype="string")
        empty["data_station"] = pd.Series(dtype="string")
        empty["match_score"] = pd.Series(dtype="float")
        empty["match_method"] = pd.Series(dtype="string")
        return empty, None, {}

    # Station symbols need true Point geometries; convert any MultiPoint or
    # other unexpected geometry to its centroid.
    if not gdf.geometry.geom_type.eq("Point").all():
        warnings.warn(f"{source}: non-Point features detected; using centroids for map station symbols.")
        gdf = gdf.copy()
        gdf.geometry = gdf.geometry.centroid

    name_field, field_scores = choose_station_name_field(gdf, excel_names)
    layer = gdf.copy()
    layer["source"] = source

    if name_field is None:
        layer["gis_station_name"] = ""
        layer["data_station"] = None
        layer["match_score"] = 0.0
        layer["match_method"] = "no_name_field"
        return layer, None, field_scores

    layer["gis_station_name"] = layer[name_field].astype("string")
    matched = layer["gis_station_name"].apply(lambda x: match_to_excel_name(x, excel_names))
    layer["data_station"] = matched.apply(lambda x: x[0])
    layer["match_score"] = matched.apply(lambda x: x[1])
    layer["match_method"] = matched.apply(lambda x: x[2])
    return layer, name_field, field_scores


def choose_working_crs(barangays: gpd.GeoDataFrame):
    if barangays.crs is None:
        raise ValueError("Barangay layer has no CRS.")

    try:
        if barangays.crs.is_geographic:
            utm = barangays.estimate_utm_crs()
            if utm is not None:
                return utm
    except Exception:
        pass
    return barangays.crs


def polygon_mask(union_geom, xx: np.ndarray, yy: np.ndarray) -> np.ndarray:
    if contains_xy is not None:
        return contains_xy(union_geom, xx, yy)

    # Older-Shapely fallback. This is slower but only used for ~48k grid cells.
    flat = [union_geom.contains(Point(x, y)) for x, y in zip(xx.ravel(), yy.ravel())]
    return np.asarray(flat, dtype=bool).reshape(xx.shape)


def idw_surface(x: np.ndarray, y: np.ndarray, z: np.ndarray, bounds, grid_size=220, power=2.0):
    """Simple inverse-distance-weighted interpolation."""
    minx, miny, maxx, maxy = bounds
    gx = np.linspace(minx, maxx, grid_size)
    gy = np.linspace(miny, maxy, grid_size)
    xx, yy = np.meshgrid(gx, gy)

    # Chunking limits temporary array size.
    flat_x = xx.ravel()
    flat_y = yy.ravel()
    out = np.empty_like(flat_x, dtype=float)
    chunk = 10000

    for i in range(0, len(flat_x), chunk):
        xs = flat_x[i:i + chunk, None]
        ys = flat_y[i:i + chunk, None]
        d2 = (xs - x[None, :]) ** 2 + (ys - y[None, :]) ** 2

        exact = d2 < 1e-12
        distances = np.sqrt(np.maximum(d2, 1e-12))
        weights = 1.0 / np.power(distances, power)
        vals = np.sum(weights * z[None, :], axis=1) / np.sum(weights, axis=1)

        # Preserve exact station values when a grid cell lands on a station.
        if exact.any():
            rows = np.where(exact.any(axis=1))[0]
            for r in rows:
                vals[r] = z[np.argmax(exact[r])]

        out[i:i + chunk] = vals

    return xx, yy, out.reshape(xx.shape)


# =============================================================================
# PEAK-HOUR MAP
# =============================================================================


def make_peak_hour_map(stats: dict) -> tuple[Path, Path, dict]:
    excel_names = list(stats["rain"].columns)
    peak_time = stats["peak_network_time"]
    peak_values = stats["rain"].loc[peak_time]

    barangays = safe_read_vector(BARANGAYS_SHP, "Barangay boundaries")
    waterways = safe_read_vector(WATERWAYS_SHP, "Waterways")
    rain_gauges = safe_read_vector(RAIN_GAUGES_SHP, "Rain gauges")
    ews = safe_read_vector(EWS_SHP, "Early Warning Systems")

    aws, aws_search_fields = filter_aws(ews)

    rg_layer, rg_name_field, rg_field_scores = prepare_station_layer(rain_gauges, "Rain Gauge", excel_names)
    aws_layer, aws_name_field, aws_field_scores = prepare_station_layer(aws, "AWS", excel_names)

    # Normalize CRS *before* concatenation. GeoPandas can reject concatenation
    # when GeoDataFrames carry different CRS metadata.
    if not aws_layer.empty and aws_layer.crs != rain_gauges.crs:
        aws_layer = aws_layer.to_crs(rain_gauges.crs)

    # Combine GIS stations, then attach rainfall observed at the network peak hour.
    stations = pd.concat([rg_layer, aws_layer], ignore_index=True)
    stations = gpd.GeoDataFrame(stations, geometry="geometry", crs=rain_gauges.crs)

    stations["rain_mm"] = stations["data_station"].map(peak_values.to_dict())

    # Reproject all map layers to a projected CRS for meaningful IDW distances.
    working_crs = choose_working_crs(barangays)
    barangays = barangays.to_crs(working_crs)
    waterways = waterways.to_crs(working_crs)
    stations = stations.to_crs(working_crs)

    # Diagnostics CSV: important for auditing fuzzy name matches.
    diagnostics_cols = [
        "source", "gis_station_name", "data_station", "match_score", "match_method", "rain_mm"
    ]
    diagnostics = stations[diagnostics_cols].copy()
    diagnostics = diagnostics.sort_values(["source", "match_method", "gis_station_name"], na_position="last")
    match_csv = OUTPUT_DIR / "RF_260814_peak_hour_station_matches.csv"
    diagnostics.to_csv(match_csv, index=False, encoding="utf-8-sig")

    matched = stations[stations["rain_mm"].notna()].copy()
    if matched.empty:
        raise ValueError(
            "No GIS station names could be matched to Excel rainfall columns. "
            f"Review {match_csv.name} and, if needed, add station-name aliases in normalize_station_name()."
        )

    # Avoid overweighting one Excel station if it appears in both GIS datasets.
    # Rain Gauge locations are preferred over AWS locations for interpolation.
    interp = matched.copy()
    interp["source_priority"] = interp["source"].map({"Rain Gauge": 0, "AWS": 1}).fillna(2)
    interp = interp.sort_values("source_priority").drop_duplicates(subset="data_station", keep="first")

    x = interp.geometry.x.to_numpy(float)
    y = interp.geometry.y.to_numpy(float)
    z = interp["rain_mm"].to_numpy(float)

    # Intensity bins; upper bound grows if an observation exceeds 60 mm/hr.
    upper = max(60.0, math.ceil(max(float(np.nanmax(z)), 30.0) / 10.0) * 10.0)
    levels = [0.0, 2.5, 7.5, 15.0, 30.0, upper]
    colors = ["lightskyblue", "orange", "red", "darkred", "violet"]
    cmap = ListedColormap(colors)
    norm = BoundaryNorm(levels, cmap.N)

    union_geom = barangays.geometry.union_all() if hasattr(barangays.geometry, "union_all") else barangays.unary_union
    bounds = union_geom.bounds
    xx, yy, zz = idw_surface(x, y, z, bounds, grid_size=IDW_GRID_SIZE, power=IDW_POWER)
    inside = polygon_mask(union_geom, xx, yy)
    zz = np.ma.masked_where(~inside, zz)

    fig, ax = plt.subplots(figsize=(11, 11))

    ax.contourf(xx, yy, zz, levels=levels, cmap=cmap, norm=norm, alpha=0.55, antialiased=True)
    barangays.plot(ax=ax, facecolor="none", edgecolor="dimgray", linewidth=0.6, zorder=3)
    waterways.plot(ax=ax, color="royalblue", linewidth=0.7, alpha=0.75, zorder=4)

    # Plot matched station measurements with source-specific symbols.
    rg = matched[matched["source"] == "Rain Gauge"]
    aw = matched[matched["source"] == "AWS"]
    unmatched = stations[stations["rain_mm"].isna()].copy()

    if not rg.empty:
        ax.scatter(
            rg.geometry.x, rg.geometry.y,
            c=rg["rain_mm"], cmap=cmap, norm=norm,
            s=48, marker="o", edgecolors="black", linewidths=0.55, zorder=6,
        )
    if not aw.empty:
        ax.scatter(
            aw.geometry.x, aw.geometry.y,
            c=aw["rain_mm"], cmap=cmap, norm=norm,
            s=65, marker="^", edgecolors="black", linewidths=0.65, zorder=7,
        )

    # Show AWS locations that were identified in EWS but could not be linked to a
    # rainfall column, without assigning them a fabricated rainfall value.
    unmatched_aws = unmatched[unmatched["source"] == "AWS"]
    if not unmatched_aws.empty:
        ax.scatter(
            unmatched_aws.geometry.x, unmatched_aws.geometry.y,
            facecolors="none", edgecolors="gray", s=55, marker="^", linewidths=0.8,
            zorder=5,
        )

    # Label the top five station values at the network peak hour.
    top = matched.nlargest(min(5, len(matched)), "rain_mm")
    for _, row in top.iterrows():
        ax.annotate(
            f"{row['data_station']}\n{row['rain_mm']:.1f} mm",
            xy=(row.geometry.x, row.geometry.y),
            xytext=(5, 5), textcoords="offset points",
            fontsize=7.3, zorder=8,
            bbox=dict(boxstyle="round,pad=0.2", facecolor="white", alpha=0.78, edgecolor="gray"),
        )

    intensity_handles = [
        Patch(facecolor="lightskyblue", edgecolor="none", label="< 2.5 mm/hr"),
        Patch(facecolor="orange", edgecolor="none", label="Moderate: 2.5-<7.5"),
        Patch(facecolor="red", edgecolor="none", label="Heavy: 7.5-<15"),
        Patch(facecolor="darkred", edgecolor="none", label="Intense: 15-<30"),
        Patch(facecolor="violet", edgecolor="none", label="Torrential: >=30"),
    ]
    source_handles = [
        Line2D([0], [0], marker="o", linestyle="none", markerfacecolor="white", markeredgecolor="black", label="Rain Gauge"),
        Line2D([0], [0], marker="^", linestyle="none", markerfacecolor="white", markeredgecolor="black", label="AWS"),
        Line2D([0], [0], color="royalblue", lw=1.2, label="Waterway"),
    ]

    leg1 = ax.legend(handles=intensity_handles, title="Rainfall intensity", loc="upper left", fontsize=8)
    ax.add_artist(leg1)
    ax.legend(handles=source_handles, title="Map features", loc="lower left", fontsize=8)

    ax.set_title(
        "Quezon City Rainfall Intensity at Network Peak Hour\n"
        f"{peak_time.strftime('%d %B %Y, %I:%M %p')} PHT | Network mean = {stats['peak_network_mean']:.2f} mm",
        fontsize=13,
    )
    ax.set_axis_off()
    ax.set_aspect("equal")
    fig.tight_layout()

    out = OUTPUT_DIR / "RF_260814_peak_hour_rainfall_map.png"
    fig.savefig(out, dpi=240, bbox_inches="tight")
    plt.close(fig)

    map_meta = {
        "aws_count": len(aws),
        "aws_search_fields": aws_search_fields,
        "rg_name_field": rg_name_field,
        "aws_name_field": aws_name_field,
        "rg_field_scores": rg_field_scores,
        "aws_field_scores": aws_field_scores,
        "gis_station_count": len(stations),
        "gis_matched_count": int(stations["rain_mm"].notna().sum()),
        "gis_unmatched_count": int(stations["rain_mm"].isna().sum()),
        "unique_interp_count": len(interp),
        "working_crs": str(working_crs),
    }
    return out, match_csv, map_meta


# =============================================================================
# TEXT REPORT
# =============================================================================


def write_report(stats: dict, duplicate_groups: dict[str, list[str]], map_meta: dict | None = None) -> Path:
    out = OUTPUT_DIR / "RF_260814_report_0000-1200.txt"
    rain = stats["rain"]
    hourly_n = stats["hourly_n"]
    station_obs_n = stats["station_obs_n"]

    expected_start = ANALYSIS_DATE + pd.Timedelta(START_TIME + ":00")
    expected_end = ANALYSIS_DATE + pd.Timedelta(END_TIME + ":00")
    expected_times = pd.date_range(expected_start, expected_end, freq="h", inclusive="both" if END_INCLUSIVE else "left")
    actual_times = pd.DatetimeIndex(rain.index)
    missing_times = expected_times.difference(actual_times)

    peak_hour_top = stats["peak_hour_values"].dropna().head(10)

    lines: list[str] = []
    lines.append("QUEZON CITY RAINFALL REPORT")
    lines.append("14 August 2026 | 00:00-12:00 PHT")
    lines.append("=" * 72)
    lines.append("")
    lines.append("ANALYSIS BASIS")
    lines.append(f"- Workbook: {INPUT_XLSX}")
    lines.append(f"- Timestamps included: {actual_times.min():%Y-%m-%d %H:%M} to {actual_times.max():%Y-%m-%d %H:%M}")
    lines.append(f"- Hourly rows used: {len(rain)}")
    lines.append(f"- Unique rainfall stations after duplicate-column consolidation: {rain.shape[1]}")
    lines.append(f"- Valid stations per hour: min={int(hourly_n.min())}, max={int(hourly_n.max())}, mean={hourly_n.mean():.1f}")
    if len(missing_times):
        lines.append("- WARNING: Missing expected timestamps: " + ", ".join(t.strftime("%H:%M") for t in missing_times))
    else:
        lines.append("- All expected hourly timestamps in the configured period are present.")

    if duplicate_groups:
        lines.append(f"- Duplicate station header groups consolidated: {len(duplicate_groups)}")
        for base, originals in sorted(duplicate_groups.items()):
            lines.append(f"    * {base}: {', '.join(originals)}")
    else:
        lines.append("- No duplicate station header groups were detected.")

    lines.append("")
    lines.append("ANSWERS")
    lines.append(
        f"1. Mean accumulated rainfall from 00:00 to 12:00: {stats['mean_accum']:.2f} mm."
    )
    lines.append(
        "   Definition: rainfall was accumulated separately for each unique station over the selected timestamps, "
        "then the ordinary mean of those station accumulations was taken."
    )
    lines.append(
        f"2. Peak network-mean hourly rainfall: {stats['peak_network_mean']:.2f} mm at "
        f"{stats['peak_network_time']:%Y-%m-%d %H:%M} PHT."
    )
    lines.append(
        f"3. Highest individual hourly station observation: {stats['peak_station_value']:.2f} mm at "
        f"{stats['peak_station']} on {stats['peak_station_time']:%Y-%m-%d %H:%M} PHT."
    )
    lines.append(
        f"4. Highest station accumulated rainfall in the period: {stats['max_accum_value']:.2f} mm at "
        f"{stats['max_accum_station']}."
    )

    lines.append("")
    lines.append("TOP STATIONS AT THE NETWORK PEAK HOUR")
    for rank, (station, value) in enumerate(peak_hour_top.items(), start=1):
        lines.append(f"{rank:>2}. {station}: {value:.2f} mm")

    lines.append("")
    lines.append("DATA COMPLETENESS BY STATION")
    full_n = len(rain)
    incomplete = station_obs_n[station_obs_n < full_n].sort_values()
    if incomplete.empty:
        lines.append(f"- All stations have {full_n}/{full_n} valid hourly observations in the selected period.")
    else:
        lines.append(f"- {len(incomplete)} station(s) have missing observations:")
        for station, n in incomplete.items():
            lines.append(f"    * {station}: {int(n)}/{full_n} valid hourly observations")

    if map_meta is not None:
        lines.append("")
        lines.append("GIS / PEAK-HOUR MAP DIAGNOSTICS")
        lines.append(f"- AWS features identified in EWS shapefile: {map_meta['aws_count']}")
        lines.append(f"- Rain-gauge name field selected: {map_meta['rg_name_field']}")
        lines.append(f"- AWS name field selected: {map_meta['aws_name_field']}")
        lines.append(f"- GIS station features considered: {map_meta['gis_station_count']}")
        lines.append(f"- GIS features matched to an Excel rainfall station: {map_meta['gis_matched_count']}")
        lines.append(f"- GIS features not matched to rainfall data: {map_meta['gis_unmatched_count']}")
        lines.append(f"- Unique station measurements used for IDW interpolation: {map_meta['unique_interp_count']}")
        lines.append(f"- Working map CRS: {map_meta['working_crs']}")
        lines.append("- Full GIS matching details: RF_260814_peak_hour_station_matches.csv")

    lines.append("")
    lines.append("INTERPRETATION NOTE")
    lines.append(
        "The citywide accumulated figure is a network mean, not the rainfall total at every location. "
        "Localized station totals and hourly peaks can be substantially higher or lower. The peak-hour map uses "
        "IDW interpolation as a visualization aid; station observations remain the authoritative measured values."
    )

    out.write_text("\n".join(lines), encoding="utf-8")
    return out


# =============================================================================
# MAIN
# =============================================================================


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Loading rainfall workbook...")
    df, duplicate_groups = load_rainfall()
    period = select_period(df)
    stats = calculate_statistics(period)

    print("Creating hourly mean rainfall plot...")
    hourly_plot = plot_hourly_mean(stats)

    map_meta = None
    map_path = None
    match_csv = None
    try:
        print("Loading GIS layers and creating peak-hour rainfall map...")
        map_path, match_csv, map_meta = make_peak_hour_map(stats)
    except Exception as exc:
        # Still produce the rainfall report and bar plot if GIS matching fails.
        warnings.warn(f"Peak-hour map could not be generated: {exc}")
        map_meta = None

    print("Writing rainfall report...")
    report = write_report(stats, duplicate_groups, map_meta=map_meta)

    print("\nAnalysis complete.")
    print(f"Hourly plot: {hourly_plot}")
    print(f"Report:      {report}")
    if map_path is not None:
        print(f"Peak map:    {map_path}")
    if match_csv is not None:
        print(f"GIS matches: {match_csv}")

    print("\nKey results:")
    print(f"Mean accumulated rainfall: {stats['mean_accum']:.2f} mm")
    print(
        f"Network peak hour: {stats['peak_network_time']:%Y-%m-%d %H:%M} PHT "
        f"({stats['peak_network_mean']:.2f} mm mean)"
    )
    print(
        f"Highest individual hourly observation: {stats['peak_station_value']:.2f} mm at "
        f"{stats['peak_station']} ({stats['peak_station_time']:%Y-%m-%d %H:%M} PHT)"
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        raise
