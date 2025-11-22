from __future__ import annotations

from pathlib import Path
from typing import Callable, Dict, Optional, Tuple
import pandas as pd
import numpy as np
import re

# --------------------------------------------------------------------------------------
# Filename & sheet-name helpers
# --------------------------------------------------------------------------------------

_FORBIDDEN = r'<>:"/\\|?*\x00-\x1F'

def safe_filename(name: str) -> str:
    """Sanitize to a safe filename (Windows/macOS friendly)."""
    if name is None:
        return "unnamed"
    name = re.sub(f"[{_FORBIDDEN}]", "_", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(" .")
    return name or "unnamed"

def safe_sheet_name(name: str) -> str:
    """Excel sheet name: max 31 chars, no []:*?/\\ and cannot end with '"""
    if name is None:
        name = "Sheet"
    name = re.sub(r"[][:*?/\\]", "_", str(name))
    name = name.strip("'")
    return name[:31] if len(name) > 31 else name

# --------------------------------------------------------------------------------------
# Core logic
# --------------------------------------------------------------------------------------

def _ensure_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    lower = {c.lower(): c for c in df.columns}
    ren = {}
    for want in ["meter_id", "meter_name", "date", "reading"]:
        if want not in df.columns and want in lower:
            ren[lower[want]] = want
    if ren:
        df.rename(columns=ren, inplace=True)

    if "meter_id" not in df.columns:
        if "meter_name" in df.columns:
            df["meter_id"] = df["meter_name"].astype(str)
        else:
            df["meter_id"] = "m" + df.reset_index().index.astype(str)

    if not np.issubdtype(df["date"].dtype, np.datetime64):
        df["date"] = pd.to_datetime(df["date"], errors="coerce", utc=False)

    df["reading"] = pd.to_numeric(df["reading"], errors="coerce")
    df.sort_values(["meter_id", "date"], inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df

def _compute_periods(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["prev_date"] = df.groupby("meter_id")["date"].shift(1)
    df["prev_reading"] = df.groupby("meter_id")["reading"].shift(1)
    df["days"] = (df["date"] - df["prev_date"]).dt.days
    df["consumption"] = df["reading"] - df["prev_reading"]
    df["daily_rate"] = np.where(df["days"] > 0, df["consumption"] / df["days"], np.nan)
    df["annualized_consumption"] = df["daily_rate"] * 365.0
    df["Quelle"] = "gemessen"
    return df

def _interp_reading_for_date(sub: pd.DataFrame, target_date: pd.Timestamp) -> Tuple[float, float]:
    sub = sub.sort_values("date")
    match = sub.loc[sub["date"] == target_date]
    if not match.empty:
        row = match.iloc[0]
        return float(row["reading"]), float(row["daily_rate"] if pd.notna(row["daily_rate"]) else 0.0)

    before = sub[sub["date"] <= target_date].tail(1)
    after = sub[sub["date"] >= target_date].head(1)

    if not before.empty and not after.empty and before.iloc[0]["date"] != after.iloc[0]["date"]:
        d0 = before.iloc[0]["date"]
        r0 = before.iloc[0]["reading"]
        d1 = after.iloc[0]["date"]
        r1 = after.iloc[0]["reading"]
        total_days = (d1 - d0).days
        if total_days > 0:
            frac = (target_date - d0).days / total_days
            reading_est = r0 + frac * (r1 - r0)
            period = after.iloc[0]
            daily_rate = float(period["daily_rate"]) if pd.notna(period["daily_rate"]) else float((r1 - r0) / total_days)
            return float(reading_est), daily_rate

    if not before.empty:
        row = before.iloc[0]
        rate = float(row["daily_rate"]) if pd.notna(row["daily_rate"]) else 0.0
        delta_days = (target_date - row["date"]).days
        return float(row["reading"] + rate * delta_days), rate

    if not after.empty:
        row = after.iloc[0]
        rate = float(row["daily_rate"]) if pd.notna(row["daily_rate"]) else 0.0
        delta_days = (row["date"] - target_date).days
        return float(row["reading"] - rate * delta_days), rate

    return (np.nan, np.nan)

def _normalized_rows(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    if df["date"].notna().any():
        year_min = int(df["date"].min().year)
        year_max = int(df["date"].max().year)
    else:
        return pd.DataFrame(columns=df.columns)

    for meter_id, sub in df.groupby("meter_id"):
        meter_name = sub["meter_name"].iloc[0] if "meter_name" in sub.columns else meter_id
        unit = sub["unit"].iloc[0] if "unit" in sub.columns else None

        for y in range(year_min, year_max + 1):
            jan1 = pd.Timestamp(year=y, month=1, day=1)
            dec31 = pd.Timestamp(year=y, month=12, day=31)

            for target in (jan1, dec31):
                reading_est, rate_est = _interp_reading_for_date(sub, target)
                if pd.isna(reading_est):
                    continue
                rows.append({
                    "meter_id": meter_id,
                    "meter_name": meter_name,
                    "date": target,
                    "reading": reading_est,
                    "prev_date": pd.NaT,
                    "prev_reading": np.nan,
                    "days": np.nan,
                    "consumption": np.nan,
                    "daily_rate": rate_est if pd.notna(rate_est) else np.nan,
                    "annualized_consumption": (rate_est * 365.0) if pd.notna(rate_est) else np.nan,
                    "Quelle": "ermittelt",
                    **({"unit": unit} if unit is not None else {}),
                })

    if not rows:
        return pd.DataFrame(columns=df.columns)
    norm = pd.DataFrame(rows)
    norm["date"] = pd.to_datetime(norm["date"])
    return norm

def build_consumption_sheet(df_readings: pd.DataFrame) -> pd.DataFrame:
    base = _ensure_cols(df_readings)
    with_periods = _compute_periods(base)
    normalized = _normalized_rows(with_periods)
    out = pd.concat([with_periods, normalized], ignore_index=True, sort=False)
    out.sort_values(["meter_id", "date", "Quelle"], inplace=True)

    cols = ["meter_id", "meter_name", "date", "reading", "days", "consumption",
            "daily_rate", "annualized_consumption", "Quelle"]
    if "unit" in out.columns:
        cols.insert(3, "unit")
    extra = [c for c in out.columns if c not in cols]
    out = out[cols + extra]
    return out

# --------------------------------------------------------------------------------------
# Writers (with optional custom sanitizers)
# --------------------------------------------------------------------------------------

def write_folder_workbook(xlsx_path: Path | str,
                          main_sheet_name: str,
                          df_main: pd.DataFrame,
                          df_readings: pd.DataFrame,
                          consumption_sheet_name: str = "Verbrauch+Norm",
                          sanitize_filename: Optional[Callable[[str], str]] = None,
                          sanitize_sheet: Optional[Callable[[str], str]] = None) -> None:
    sanitize_filename = sanitize_filename or safe_filename
    sanitize_sheet = sanitize_sheet or safe_sheet_name

    xlsx_path = Path(xlsx_path)
    xlsx_path = xlsx_path.with_name(sanitize_filename(xlsx_path.name))
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    df_cons = build_consumption_sheet(df_readings)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_main.to_excel(writer, sheet_name=sanitize_sheet(main_sheet_name), index=False)
        df_cons.to_excel(writer, sheet_name=sanitize_sheet(consumption_sheet_name), index=False)

def write_master_workbook(xlsx_path: Path | str,
                          per_folder_frames: Dict[str, pd.DataFrame],
                          master_sheet_name: str = "Alle Zähler",
                          sanitize_filename: Optional[Callable[[str], str]] = None,
                          sanitize_sheet: Optional[Callable[[str], str]] = None) -> None:
    sanitize_filename = sanitize_filename or safe_filename
    sanitize_sheet = sanitize_sheet or safe_sheet_name

    xlsx_path = Path(xlsx_path)
    xlsx_path = xlsx_path.with_name(sanitize_filename(xlsx_path.name))
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    frames = []
    for folder, df_readings in per_folder_frames.items():
        df_cons = build_consumption_sheet(df_readings)
        df_cons.insert(0, "folder", folder)
        frames.append(df_cons)

    master = pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        master.to_excel(writer, sheet_name=sanitize_sheet(master_sheet_name), index=False)

