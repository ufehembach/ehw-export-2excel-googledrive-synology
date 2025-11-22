import pandas as pd
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
import re as _re

def build_yearly_view(df):
    df2 = df.copy()
    df2["Date"] = pd.to_datetime(df2["Date_Full"], errors="coerce")
    df2 = df2.sort_values(["CounterId", "Date"])
    df2["Year"] = df2["Date"].dt.year

    yearly = []
    for cid, group in df2.groupby("CounterId"):
        for year, g2 in group.groupby("Year"):
            year_end = pd.Timestamp(year=year, month=12, day=31)
            g_before = g2[g2["Date"] <= year_end]

            if not g_before.empty:
                row = g_before.iloc[-1]
            else:
                g_after = group[
                    (group["Date"] > year_end) &
                    (group["Date"] <= year_end + pd.Timedelta(days=15))
                ]
                if g_after.empty:
                    continue
                row = g_after.iloc[0]

            yearly.append({
                "Object": row["Object"],
                "Room": row["Room"],
                "CounterName": row["CounterName"],
                "CounterId": cid,
                "Year": year,
                "Date": row["Date"],
                "Value": row["Value_Num"],
                "CounterType": row["CounterType"],
                "CounterUnit": row["CounterUnit"],
            })

    df_year = pd.DataFrame(yearly).sort_values(["CounterId", "Year"])
    # --- Add delta/prev with reset detection ---
    prev_vals = []
    prev_dates = []
    deltas = []

    last_value = {}
    last_date = {}

    for idx, row in df_year.iterrows():
        key = row["CounterId"]
        current_val = row["Value"]
        current_date = row["Date"]

        if key not in last_value:
            prev_vals.append(None)
            prev_dates.append(None)
            deltas.append(None)
            last_value[key] = current_val
            last_date[key] = current_date
            continue

        previous_val = last_value[key]
        previous_date = last_date[key]

        if current_val < previous_val:
            delta = current_val
            prev_vals.append(None)
            prev_dates.append(None)
        else:
            delta = current_val - previous_val
            prev_vals.append(previous_val)
            prev_dates.append(previous_date)

        deltas.append(delta)

        last_value[key] = current_val
        last_date[key] = current_date

    df_year["PrevValue"] = prev_vals
    df_year["PrevDate"] = prev_dates
    df_year["Delta"] = deltas

    # --- Add DeltaPerDay ---
    delta_per_day = []
    for idx, row in df_year.iterrows():
        if row["PrevDate"] is None or row["Delta"] is None:
            delta_per_day.append(None)
        else:
            days = (row["Date"] - row["PrevDate"]).days
            if days > 0:
                delta_per_day.append(row["Delta"] / days)
            else:
                delta_per_day.append(None)
    df_year["DeltaPerDay"] = delta_per_day

    # --- Add Days ---
    days_list = []
    for idx, row in df_year.iterrows():
        if row["PrevDate"] is None:
            days_list.append(None)
        else:
            days_list.append((row["Date"] - row["PrevDate"]).days)
    df_year["Days"] = days_list

    return df_year


def build_monthly_view(df):
    df2 = df.copy()
    df2["Date"] = pd.to_datetime(df2["Date_Full"], errors="coerce")
    df2 = df2.sort_values(["CounterId", "Date"])
    df2["YearMonth"] = df2["Date"].dt.to_period("M")

    monthly = []
    for cid, group in df2.groupby("CounterId"):
        min_month = group["YearMonth"].min()
        max_month = group["YearMonth"].max()

        for ym in pd.period_range(min_month, max_month, freq="M"):
            month_end = ym.to_timestamp(how="end")
            g_before = group[group["Date"] <= month_end]

            if not g_before.empty:
                row = g_before.iloc[-1]
            else:
                g_after = group[
                    (group["Date"] > month_end) &
                    (group["Date"] <= month_end + pd.Timedelta(days=10))
                ]
                if g_after.empty:
                    continue
                row = g_after.iloc[0]

            monthly.append({
                "Object": row["Object"],
                "Room": row["Room"],
                "CounterName": row["CounterName"],
                "CounterId": cid,
                "YearMonth": str(ym),
                "Date": row["Date"],
                "Value": row["Value_Num"],
                "CounterType": row["CounterType"],
                "CounterUnit": row["CounterUnit"],
            })

    df_month = pd.DataFrame(monthly).sort_values(["CounterId", "YearMonth"])

    # --- Add delta/prev with reset detection ---
    prev_vals = []
    prev_dates = []
    deltas = []

    last_value = {}
    last_date = {}

    for idx, row in df_month.iterrows():
        key = row["CounterId"]
        current_val = row["Value"]
        current_date = row["Date"]

        if key not in last_value:
            # first entry → no delta
            prev_vals.append(None)
            prev_dates.append(None)
            deltas.append(None)
            last_value[key] = current_val
            last_date[key] = current_date
            continue

        previous_val = last_value[key]
        previous_date = last_date[key]

        if current_val < previous_val:
            # reset detected
            delta = current_val
            prev_vals.append(None)
            prev_dates.append(None)
        else:
            delta = current_val - previous_val
            prev_vals.append(previous_val)
            prev_dates.append(previous_date)

        deltas.append(delta)

        last_value[key] = current_val
        last_date[key] = current_date

    df_month["PrevValue"] = prev_vals
    df_month["PrevDate"] = prev_dates
    df_month["Delta"] = deltas

    # --- Add DeltaPerDay (Verbrauch pro Tag) ---
    delta_per_day = []
    for idx, row in df_month.iterrows():
        if row["PrevDate"] is None or row["Delta"] is None:
            delta_per_day.append(None)
        else:
            days = (row["Date"] - row["PrevDate"]).days
            if days > 0:
                delta_per_day.append(row["Delta"] / days)
            else:
                delta_per_day.append(None)
    df_month["DeltaPerDay"] = delta_per_day

    # --- Add Days (Anzahl Tage zwischen Ablesungen) ---
    days_list = []
    for idx, row in df_month.iterrows():
        if row["PrevDate"] is None:
            days_list.append(None)
        else:
            days_list.append((row["Date"] - row["PrevDate"]).days)
    df_month["Days"] = days_list

    return df_month

def extract_unit(counter_name: str) -> str:
    """
    Extracts the unit (Wohnungseinheit), e.g. DBMP.EG or H1.Whg2
    from CounterName like 'DBMP.EG.Wasser-Küche'.
    """
    if not isinstance(counter_name, str):
        return ""
    parts = counter_name.split(".")
    if len(parts) >= 2:
        return ".".join(parts[:2])
    return parts[0]


def extract_art(counter_type: str, counter_name: str) -> str:
    """
    Detects water/heating/electricity from CounterType or CounterName.
    """
    text = f"{counter_type} {counter_name}".lower()
    if "wasser" in text or "water" in text:
        return "wasser"
    if "wärme" in text or "waerme" in text or "heat" in text:
        return "wärme"
    if "strom" in text or "electric" in text:
        return "strom"
    return ""

def build_summary_table(df_month):
    """
    Builds a clean pivoted summary:
    YearMonth | wasser | wärme | strom | Einheit
    """
    df2 = df_month.copy()
    df2["Ablesung"] = df2["Date"].dt.strftime("%Y.%m")
    df2["Einheit"] = df2["CounterName"].apply(extract_unit)
    df2["Art"] = df2.apply(lambda r: extract_art(r["CounterType"], r["CounterName"]), axis=1)

    summary = (
        df2.pivot_table(
            index="Ablesung",
            columns="Art",
            values="Value",
            aggfunc="max"
        )
        .reset_index()
        .rename_axis(None, axis=1)
    )

    einheit_map = df2.groupby("Ablesung")["Einheit"].first()
    summary["Einheit"] = summary["Ablesung"].map(einheit_map)

    summary = summary.sort_values("Ablesung")
    return summary

def add_delta_columns(df):
    """
    Adds PrevValue, PrevDate, Delta, DeltaPerDay, Days to raw df.
    Reset detection included.
    """
    df2 = df.copy()
    df2["Date_Full"] = pd.to_datetime(df2["Date_Full"], errors="coerce")
    df2 = df2.sort_values(["CounterId", "Date_Full"])

    prev_vals = []
    prev_dates = []
    deltas = []
    last_value = {}
    last_date = {}

    for idx, row in df2.iterrows():
        key = row["CounterId"]
        current_val = row.get("Value_Num", None)
        current_date = row["Date_Full"]

        if key not in last_value:
            prev_vals.append(None)
            prev_dates.append(None)
            deltas.append(None)
            last_value[key] = current_val
            last_date[key] = current_date
            continue

        previous_val = last_value[key]
        previous_date = last_date[key]

        if current_val is None or previous_val is None:
            prev_vals.append(None)
            prev_dates.append(None)
            deltas.append(None)
        elif current_val < previous_val:
            prev_vals.append(None)
            prev_dates.append(None)
            deltas.append(current_val)
        else:
            prev_vals.append(previous_val)
            prev_dates.append(previous_date)
            deltas.append(current_val - previous_val)

        last_value[key] = current_val
        last_date[key] = current_date

    df2["PrevValue"] = prev_vals
    df2["PrevDate"] = prev_dates
    df2["Delta"] = deltas

    delta_per_day = []
    days_list = []
    for idx, row in df2.iterrows():
        if row["PrevDate"] is None or row["Delta"] is None:
            delta_per_day.append(None)
            days_list.append(None)
        else:
            days = (row["Date_Full"] - row["PrevDate"]).days
            days_list.append(days if days > 0 else None)
            if days and days > 0:
                delta_per_day.append(row["Delta"] / days)
            else:
                delta_per_day.append(None)

    df2["DeltaPerDay"] = delta_per_day
    df2["Days"] = days_list

    return df2

def add_table_to_sheet(sheet_name, table_style):
    if sheet_name not in wb.sheetnames:
        return
    wsx = wb[sheet_name]
    if wsx.max_row < 2 or wsx.max_column < 1:
        return

    last_col_letter = get_column_letter(wsx.max_column)
    last_row = wsx.max_row
    table_range = f"A1:{last_col_letter}{last_row}"

    if sheet_name == "Zählerdaten_Jahr":
        table_name = "tblehwJahr"
    elif sheet_name == "Zählerdaten_Monat":
        table_name = "tblehwMonat"
    else:
        table_name = "TblData"

    t = Table(displayName=table_name, ref=table_range)
    t.tableStyleInfo = TableStyleInfo(
        name=table_style,
        showRowStripes=True,
        showColumnStripes=False
    )
    wsx.add_table(t)
