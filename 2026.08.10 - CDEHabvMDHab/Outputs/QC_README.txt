QC OUTPUT GUIDE
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
