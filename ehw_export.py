#!/usr/bin/env python3
import sys
import subprocess
import importlib
import json
import pandas as pd
from pathlib import Path
import shutil
import pwd
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
import re
from ehw_transform import build_yearly_view, build_monthly_view

import ehw_fix_Images

# --- Debug flag ---
DEBUG = bool(os.getenv("EHW_DEBUG"))
DEBUG = True

# Image management settings
IMG_MODE = os.getenv("EHW_IMG_MODE", "copy").lower()  # copy | symlink
PRUNE = bool(os.getenv("EHW_PRUNE"))

CONFIG_FILE = "./ehw_export.conf.json"
MAX_XLSX_FILES = 10

# === Automatische Paketinstallation ===
def ensure_package(pkg):
    try:
        importlib.import_module(pkg)
    except ImportError:
        print(f"[INFO] Installiere fehlendes Paket: {pkg}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

for package in ["pandas", "openpyxl"]:
    ensure_package(package)

# --- Datum & Werte Parsing ---
def parse_date_variants(date_str):
    if not date_str:
        return "", "", "", ""
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", ""))
        if dt.hour == 0 and dt.minute == 0:
            orig = dt.strftime("%d.%m.%Y")
        else:
            orig = dt.strftime("%d.%m.%Y %H:%M")
        return orig, dt.strftime("%Y"), dt.strftime("%Y-%m"), dt.strftime("%Y-%m-%d")
    except Exception:
        return date_str, "", "", ""

def parse_value_numeric(val):
    if val is None:
        return None
    try:
        num_str = re.sub(r"[^0-9,.\-]", "", str(val)).replace(",", ".")
        return float(num_str) if num_str else None
    except ValueError:
        return None

import re as _re

def safe_name(s: str) -> str:
    """Sanitize names for filesystem paths."""
    if s is None:
        return "unknown"
    return _re.sub(r"[\\/|:*?\"<>]", "_", str(s)).strip()

def resolve_room_and_object(counter: dict, room_map: dict) -> tuple[str | None, str | None]:
    """
    Determine the room and object names for a given counter.
    Returns (room_name, object_name).
    """
    room_id = counter.get("roomId")
    room_name = room_map.get(room_id)
    cn = (counter.get("counterName") or "").strip()

    # Helper to extract prefix from a name using . or -
    def extract_prefix(name: str) -> str:
        for delim in ('.', '-'):
            if delim in name:
                return name.split(delim, 1)[0]
        return name

    if room_name:
        object_name = extract_prefix(room_name)
    else:
        object_name = extract_prefix(cn) if cn else None

    return room_name, object_name

# --- Canonical index builder for hidden store ---
import os as _os
import re as _re2

# Build an index of images in the hidden canonical store, supporting multiple layouts:
# 1) flat:   <root>/<obj_uuid>_<room_uuid>_<filename>
# 2) nested: <root>/<obj_uuid>/<room_uuid>/<filename>
# 3) minimal (fallback): <root>/<filename>
CANON_FLAT_RE = _re2.compile(r"^(?P<obj>[0-9a-f\-]{8,})_(?P<room>[0-9a-f\-]{8,})_(?P<file>.+)$", _re2.IGNORECASE)

def build_canonical_index(canonical_root: Path) -> dict:
    idx = {
        "by_tuple": {},      # (obj_uuid, room_uuid, file) -> Path
        "by_room_file": {},  # (room_uuid, file) -> list[Path]
        "by_file": {},       # file -> list[Path]
    }
    if not canonical_root.exists():
        return idx
    for dirpath, _dirs, files in _os.walk(canonical_root):
        dpath = Path(dirpath)
        # Case 2 and 3: nested or minimal
        parts = dpath.relative_to(canonical_root).parts
        obj_uuid = parts[0] if len(parts) >= 1 and parts[0] else None
        room_uuid = parts[1] if len(parts) >= 2 and parts[1] else None
        for fn in files:
            p = dpath / fn
            m = CANON_FLAT_RE.match(fn)
            if m:
                o = m.group("obj"); r = m.group("room"); f = m.group("file")
                idx["by_tuple"][(o, r, f)] = p
                idx["by_room_file"].setdefault((r, f), []).append(p)
                idx["by_file"].setdefault(f, []).append(p)
                continue
            # nested case
            f = fn
            if obj_uuid and room_uuid:
                idx["by_tuple"][(obj_uuid, room_uuid, f)] = p
                idx["by_room_file"].setdefault((room_uuid, f), []).append(p)
            # minimal fallback
            idx["by_file"].setdefault(f, []).append(p)
    return idx

from pathlib import Path as _Path
import shutil as _shutil

def copy_canonical_image(
    local_image_file_name: str | None,
    canonical_root: Path,
    object_uuid: str | None,
    room_uuid: str | None,
    object_name: str | None,
    room_name: str | None,
    visible_root: Path,
    verbose: bool = False,
    canon_idx: dict | None = None
) -> Path | None:
    """Copy or symlink from <target>/.<folder>/<obj>_<room>_<file> to <target>/<folder>/<Object>/<Room>/<file>.
    Returns dest path as Path or None if source doesn't exist or no filename.
    """
    if not local_image_file_name or not object_uuid or not room_uuid:
        return None
    file_name = _Path(local_image_file_name).name
    src = None
    if canon_idx:
        # 1) exact tuple
        src = canon_idx["by_tuple"].get((object_uuid, room_uuid, file_name))
        # 2) by (room_uuid, file)
        if not src:
            cands = canon_idx["by_room_file"].get((room_uuid, file_name)) or []
            src = cands[0] if cands else None
        # 3) by file only (last resort)
        if not src:
            cands = canon_idx["by_file"].get(file_name) or []
            src = cands[0] if cands else None
    if not src:
        # Legacy flat name fallback
        canonical_name = f"{object_uuid}_{room_uuid}_{file_name}"
        src = canonical_root / canonical_name
    if verbose:
        print(f"[IMG] try src={src}")
    if not src or not Path(src).exists():
        if verbose:
            print(f"[IMG] missing: {src}")
        return None
    dest_dir = visible_root / safe_name(object_name) / safe_name(room_name)
    if verbose:
        print(f"[IMG] ensure dest_dir={dest_dir}")
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / file_name
    if not dest.exists():
        if IMG_MODE == "symlink":
            if verbose:
                print(f"[IMG] symlink {src} -> {dest}")
            dest.symlink_to(src)
        else:
            if verbose:
                print(f"[IMG] copy {src} -> {dest}")
            _shutil.copy2(src, dest)
    else:
        if verbose:
            print(f"[IMG] exists: {dest}")
    return dest
# --- Helpers ---
def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def cleanup_old_excels(folder: Path, prefix: str):
    files = sorted(folder.glob(f"{prefix}*.xlsx"), key=os.path.getmtime)
    while len(files) > MAX_XLSX_FILES:
        old_file = files.pop(0)
        try:
            old_file.unlink()
            print(f"[Cleanup] Gelöscht: {old_file}")
        except Exception as e:
            print(f"[WARN] Konnte {old_file} nicht löschen: {e}")

# --- Hauptlogik ---
ALL_ROWS = []

def process_folder(sync_dir: Path, target_base_dir: Path):
    json_path = sync_dir / f"{sync_dir.name}.json"
    if not json_path.exists():
        print(f"[WARN] JSON fehlt: {json_path}")
        return

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Map each roomId (UUID) to its cleartext name
    room_map = {}
    for room in data.get("rooms", []):
        rid = room.get("roomId")
        rname = room.get("name") or room.get("title")
        if rid:
            room_map[rid] = rname
    # canonical source and visible target roots for images
    canonical_root = target_base_dir / f".{sync_dir.name}"
    visible_images_root = target_base_dir / f"{sync_dir.name}"
    object_uuid = data.get("objectId")
    if DEBUG:
        print(f"[DBG] canonical_root={canonical_root}")
        print(f"[DBG] visible_root={visible_images_root}")
        print(f"[DBG] object_uuid={object_uuid}")
    canon_idx = build_canonical_index(canonical_root)
    if DEBUG:
        print(f"[DBG] canonical indexed: tuples={len(canon_idx['by_tuple'])} by_room={len(canon_idx['by_room_file'])} by_file={len(canon_idx['by_file'])}")

    expected_files: list[Path] = []
    rows = []
    for counter in data.get("counters", []):
        for entry in counter.get("entries", {}).get("entries", []):
            date_orig, date_year, date_yearmonth, date_full = parse_date_variants(entry.get("date"))
            value_orig = entry.get("value")
            value_num = parse_value_numeric(value_orig)
            # Resolve human-readable room and object names
            room_name, object_name = resolve_room_and_object(counter, room_map)
            dest_path: Path | None = copy_canonical_image(
                entry.get("localImageFileName"),
                canonical_root=canonical_root,
                object_uuid=object_uuid,
                room_uuid=counter.get("roomId"),
                object_name=object_name,
                room_name=room_name,
                visible_root=visible_images_root,
                verbose=DEBUG,
                canon_idx=canon_idx,
            )
            if DEBUG and not entry.get("localImageFileName"):
                print(f"[DBG] no localImageFileName for counter={counter.get('counterName')} date={entry.get('date')}")
            if dest_path:
                expected_files.append(dest_path)
                try:
                    rel = dest_path.relative_to(target_base_dir)
                except Exception:
                    rel = dest_path
                img_link_formula = f'=HYPERLINK("{rel.as_posix()}","Bild")'
            else:
                img_link_formula = ""

            # Build row dictionary (CounterUUID suppressed)
            row = {
                "Object": object_name,
                "Room": room_name,
                "CounterName": counter.get("counterName"),
                "Bild": "Bild" if dest_path else "",  # text; hyperlink applied later
                "CounterType": counter.get("counterType"),
                "CounterUnit": counter.get("counterUnit"),
                "CounterId": counter.get("counterId"),
                "RoomId": counter.get("roomId"),
                "Date_Orig": date_orig,
                "Date_Year": date_year,
                "Date_YearMonth": date_yearmonth,
                "Date_Full": date_full,
                "Value_Orig": value_orig,
                "Value_Num": value_num,
            }
            # store absolute path for post-write hyperlinking
            if dest_path:
                row["__BildAbs"] = dest_path.resolve().as_posix()
            rows.append(row)

    df = pd.DataFrame(rows)
    # --- Add delta/prev/days for raw data ---
    from ehw_transform import add_delta_columns
    df = add_delta_columns(df)
    df["_SourceFolder"] = sync_dir.name
    ALL_ROWS.append(df.copy())
    # Ensure column order and keep helper path for now
    desired_order = [
        "Object", "Room", "CounterName", "Bild",
        "CounterType", "CounterUnit", "CounterId", "RoomId",
        "Date_Orig", "Date_Year", "Date_YearMonth", "Date_Full",
        "Value_Orig", "Value_Num", "__BildAbs"
    ]
    df = df.reindex(columns=[c for c in desired_order if c in df.columns])

    # Optional prune of stale images in visible tree
    if PRUNE and visible_images_root.exists():
        keep = {p.resolve() for p in expected_files}
        removed = 0
        for dirpath, _dirs, files in _os.walk(visible_images_root):
            dpath = Path(dirpath)
            for fn in files:
                p = (dpath / fn).resolve()
                if p not in keep:
                    if DEBUG:
                        print(f"[PRUNE] remove stale: {p}")
                    try:
                        p.unlink()
                        removed += 1
                    except Exception as e:
                        print(f"[WARN] prune failed for {p}: {e}")
        if DEBUG:
            print(f"[PRUNE] removed {removed} files")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_name = f"##{sync_dir.name}-{timestamp}.xlsx"
    excel_path = target_base_dir / excel_name

    export_with_format(df, excel_path, script_name="ehw_export")
    cleanup_old_excels(target_base_dir, f"##{sync_dir.name}-")

    # Also create/overwrite a stable latest file for cross-links
    latest_path = target_base_dir / f"{sync_dir.name}.xlsx"
    try:
        shutil.copy2(excel_path, latest_path)
        print(f"[OK] Latest Excel aktualisiert: {latest_path}")
    except Exception as e:
        print(f"[WARN] Konnte Latest Excel nicht schreiben: {latest_path} -> {e}")

# --- Excel-Formatierung ---
def export_with_format(df, file_path, script_name):
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name="Zählerdaten")
        # --- Build yearly and monthly aggregated counter sheets ---
        try:
            df_year = build_yearly_view(df)
            df_year.to_excel(writer, index=False, sheet_name="Zählerdaten_Jahr")
        except Exception as e:
            print(f"[WARN] Jahres-Ansicht konnte nicht erzeugt werden: {e}")

        try:
            df_month = build_monthly_view(df)
            df_month.to_excel(writer, index=False, sheet_name="Zählerdaten_Monat")
        except Exception as e:
            print(f"[WARN] Monats-Ansicht konnte nicht erzeugt werden: {e}")


        # After write: turn Bild text into hyperlinks using __BildAbs
        ws = writer.book.active
        # Find column indices
        headers = [cell.value for cell in ws[3]]
        def col_index(name):
            return headers.index(name) + 1 if name in headers else None
        col_bild = col_index("Bild")
        col_abs = col_index("__BildAbs")
        if col_bild and col_abs:
            for r in range(4, ws.max_row + 1):
                text = ws.cell(row=r, column=col_bild).value
                path = ws.cell(row=r, column=col_abs).value
                if text and path:
                    # Use absolute file URI to avoid locale/comma/semicolon formula issues
                    uri = f"file://{path}"
                    ws.cell(row=r, column=col_bild).value = text
                    ws.cell(row=r, column=col_bild).hyperlink = uri
        # Drop helper column if present by clearing its header
        if col_abs:
            ws.delete_cols(col_abs, 1)

    wb = load_workbook(file_path)
    ws = wb.active
    # --- Re-add header info in row 1 ---
    try:
        header_info = f"{script_name} -- {file_path.parent} -- {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} -- {pwd.getpwuid(os.getuid()).pw_name} -- {os.uname().nodename}"
        ws["A1"] = header_info
    except Exception as e:
        print(f"[WARN] Header row could not be written: {e}")
    # --- Minimal Excel Table Only ---
    from openpyxl.worksheet.table import Table, TableStyleInfo

    # Determine table range using non-empty header columns
    header_cells = ws[3]
    actual_cols = 0
    for idx, cell in enumerate(header_cells, start=1):
        if cell.value not in (None, "", "__BildAbs"):
            actual_cols = idx

    if actual_cols > 0 and ws.max_row >= 3:
        last_col_letter = ws.cell(row=3, column=actual_cols).column_letter
        last_row = ws.max_row
        table_range = f"A3:{last_col_letter}{last_row}"

        # Always use fixed table name "tblEHW" for Zählerdaten
        tbl = Table(displayName="tblEHW", ref=table_range)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        tbl.tableStyleInfo = style
        ws.add_table(tbl)

    # --- Add tables for yearly and monthly sheets ---
    from openpyxl.utils import get_column_letter

    def add_table_to_sheet(sheet_name, table_style):
        if sheet_name not in wb.sheetnames:
            return
        wsx = wb[sheet_name]
        if wsx.max_row < 2 or wsx.max_column < 1:
            return

        # Determine last column/row
        last_col_letter = get_column_letter(wsx.max_column)
        last_row = wsx.max_row
        table_range = f"A1:{last_col_letter}{last_row}"

        # Assign fixed table names
        if sheet_name == "Zählerdaten_Jahr":
            table_name = "tblehwJahr"
        elif sheet_name == "Zählerdaten_Monat":
            table_name = "tblehwMonat"
        elif sheet_name == "Zählerdaten":
            table_name = "tblEHW"
        else:
            table_name = "TblData"

        t = Table(displayName=table_name, ref=table_range)
        t.tableStyleInfo = TableStyleInfo(
            name=table_style,
            showRowStripes=True,
            showColumnStripes=False
        )
        wsx.add_table(t)

    # Different colors for the tables
    add_table_to_sheet("Zählerdaten_Jahr", "TableStyleMedium4")
    add_table_to_sheet("Zählerdaten_Monat", "TableStyleMedium7")

    wb.save(file_path)
    print(f"[OK] Excel gespeichert: {file_path}")

# --- Main ---
def main():
    config = load_config(path=CONFIG_FILE)
    base_sync_dir = Path(config["source_base_dir"]).resolve()
    target_base_dir = Path(config["target_base_dir"]).resolve()
    target_base_dir.mkdir(parents=True, exist_ok=True)
    if DEBUG:
        print("[DBG] EHW_DEBUG=1 (image copy debug enabled)")
    subprocess.run(
        ["python3", "ehw_fix_Images.py", "-c", CONFIG_FILE, "--apply"],
        check=True
    )
    for folder_name in config["folders"]:
        sync_dir = base_sync_dir / folder_name
        if sync_dir.exists():
            process_folder(sync_dir, target_base_dir)
        else:
            print(f"[WARN] Ordner fehlt: {sync_dir}")

    if ALL_ROWS:
        combined = pd.concat(ALL_ROWS, ignore_index=True)

        # Remove _SourceFolder entirely
        if "_SourceFolder" in combined.columns:
            combined = combined.drop(columns=["_SourceFolder"])

        combined_path = target_base_dir / "ehw+.xlsx"
        export_with_format(combined, combined_path, script_name="ehw_export_combined")
    print(f"[OK] Kombiniertes Excel gespeichert: {combined_path}")
if __name__ == "__main__":
    main()
