#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ehw_fix_Images.py (Python 3.9 compatible)
- verarbeitet NUR Dateien mit Suffix " (slash conflict)"
- entfernt den Suffix
- legt einen kanonischen, versteckten Ablageort unter <target>/.<Folder> an
- legt die Struktur **.<Folder>/<Objekt>/<Raum>/<Datei>** an
  * Objekt/Raum werden aus dem Dateinamen (UUIDs) abgeleitet und – falls vorhanden –
    per <Folder>.json von UUID → Klarname gemappt
- (optional) benennt eine bestehende Zielstruktur einmalig in .<Folder> um (Legacy-Backup)
Konfig: ehw_export.conf.json mit keys: source_base_dir, target_base_dir, folders
Erwartet: <Folder>/<Folder>.json (z.B. H1/H1.json) nur noch optional (UUID-zu-Name ist für diesen Schritt nicht nötig).
"""
import argparse
import hashlib
import json
import os
import re
from pathlib import Path
from shutil import copy2, move
from typing import Optional, Dict

SLASH_CONFLICT_RE = re.compile(r"\s*\(slash conflict\)$")
NUMBERED_COPY_RE = re.compile(r"\s\((\d+)\)$")
NAME_KEYS = {"name", "title", "displayName", "label"}
ID_KEYS = {"id", "uuid", "uid"}

def sha256sum(path: Path, chunk: int = 1024*1024) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()

def same_file(a: Path, b: Path) -> bool:
    try:
        if a.stat().st_size != b.stat().st_size:
            return False
        return sha256sum(a) == sha256sum(b)
    except FileNotFoundError:
        return False

def load_conf(conf_path: Path) -> dict:
    with open(conf_path, "r", encoding="utf-8") as f:
        return json.load(f)

def load_folder_json(src_folder: Path) -> Optional[dict]:
    """Erwartet <folder>/<folder>.json, z.B. H1/H1.json"""
    stem = src_folder.name
    cand = src_folder / f"{stem}.json"
    if cand.exists():
        try:
            with open(cand, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None

def collect_uuid_name_map(data) -> Dict[str, str]:
    """Heuristisch UUID->Name Paare aus verschachtelten Dict/List-Strukturen sammeln."""
    out: Dict[str, str] = {}

    def visit(node):
        if isinstance(node, dict):
            id_val = None
            name_val = None
            for k in ID_KEYS:
                v = node.get(k)
                if isinstance(v, str):
                    id_val = v
                    break
            for k in NAME_KEYS:
                v = node.get(k)
                if isinstance(v, str) and v.strip():
                    name_val = v.strip()
                    break
            if id_val and name_val:
                out[id_val] = name_val
            for v in node.values():
                visit(v)
        elif isinstance(node, list):
            for v in node:
                visit(v)

    if data is not None:
        visit(data)
    return out

def ensure_legacy_dot_folder(base_target: Path, folder_name: str) -> None:
    """Benennt <base>/<folder_name> einmalig in .<folder_name> um (falls existiert)."""
    cur = base_target / folder_name
    dot = base_target / f".{folder_name}"
    if cur.exists() and cur.is_dir():
        if not dot.exists():
            cur.rename(dot)

def parse_dest_parts_from_filename(name_no_suffix: str, ext: str):
    """<uuidA>_<uuidB>_<rest> -> (uuidA, uuidB, rest+ext); sonst (None, None, name+ext)."""
    base_wo_ext = name_no_suffix
    parts = base_wo_ext.split("_")
    if len(parts) >= 3:
        obj_uuid, room_uuid = parts[0], parts[1]
        rest = "_".join(parts[2:]) + ext
        return obj_uuid, room_uuid, rest
    return None, None, base_wo_ext + ext

def safe_name(s: str) -> str:
    return re.sub(r"[\\/|:*?\"<>]", "_", s).strip()

def process_folder(src_folder: Path, dst_base: Path, mode: str, apply: bool, verbose: bool):
    # 1) alte Struktur (falls vorhanden) -> .<folder>
    ensure_legacy_dot_folder(dst_base, src_folder.name)

    # 2) UUID->Name Map laden
    js = load_folder_json(src_folder)
    uuid2name = collect_uuid_name_map(js) if js else {}

    # 3) kanonischer, versteckter Ablageort <dst>/.<foldername>
    canonical_root = dst_base / f".{src_folder.name}"
    canonical_root.mkdir(parents=True, exist_ok=True)

    entries = sorted(p for p in src_folder.iterdir() if p.is_file())
    moved = skipped = errors = 0

    for src in entries:
        name = src.name
        if not SLASH_CONFLICT_RE.search(name):
            skipped += 1
            continue

        stripped = SLASH_CONFLICT_RE.sub("", name)

        # Extension ermitteln (auch Mehrfach-Suffixe)
        ext = "".join(Path(stripped).suffixes)
        base_wo_ext = stripped[:-len(ext)] if ext and stripped.endswith(ext) else Path(stripped).stem
        if not ext:
            ext = Path(stripped).suffix

        # (1), (2), ... Varianten ignorieren, WENN die Originaldatei existiert
        # Wir prüfen auf eine Zahl in Klammern am Ende des Basisnamens (ohne Extension).
        mnum = NUMBERED_COPY_RE.search(base_wo_ext)
        if mnum:
            # Kandidat für die Originaldatei MIT dem ursprünglichen '(slash conflict)'-Suffix
            original_with_suffix = f"{base_wo_ext[:mnum.start()]}{ext} (slash conflict)"
            if (src_folder / original_with_suffix).exists():
                if verbose:
                    print(f"~ IGNORE numbered duplicate (original exists): {src.name}")
                skipped += 1
                continue

        # Ziel: versteckte Ablage unter .<Folder>/<Object>/<Room>/<Datei>
        # Objekt/Raum aus dem Dateinamen (UUIDs) ableiten und mit uuid2name mappen
        obj_uuid, room_uuid, remainder = parse_dest_parts_from_filename(base_wo_ext, ext)
        obj_label = uuid2name.get(obj_uuid, obj_uuid or "unknown-object")
        room_label = uuid2name.get(room_uuid, room_uuid or "unknown-room")
        dst_dir = canonical_root / safe_name(str(obj_label)) / safe_name(str(room_label))
        dst_dir.mkdir(parents=True, exist_ok=True)
        dst = dst_dir / remainder

        try:
            if dst.exists():
                if same_file(src, dst):
                    if verbose:
                        print(f"= SKIP (identisch): {dst}")
                    skipped += 1
                    continue
                else:
                    if verbose:
                        print(f"~ OVERWRITE: {dst}")
                    if apply:
                        copy2(src, dst)
                    moved += 1
                    continue

            if verbose and not apply:
                print(f"DRYRUN -> {dst}")
            if apply:
                if mode == "copy":
                    copy2(src, dst)
                elif mode == "move":
                    try:
                        move(str(src), str(dst))
                    except Exception:
                        copy2(src, dst)
                        try:
                            src.unlink()
                        except Exception:
                            pass
                elif mode == "symlink":
                    rel_target = os.path.relpath(src, start=dst.parent)
                    os.symlink(rel_target, dst)
                else:
                    raise ValueError(f"Unknown mode: {mode}")
            moved += 1
        except Exception as e:
            errors += 1
            print(f"! ERROR {src} -> {dst}: {e}")

    if verbose:
        print(f"Summary {src_folder.name} -> .{src_folder.name}: moved={moved} skipped={skipped} errors={errors}")
    return moved, skipped, errors

def main():
    ap = argparse.ArgumentParser(description="Fix '(slash conflict)' Bilder und neue Objekt/Raum-Struktur bauen.")
    ap.add_argument("-c", "--config", default="ehw_export.conf.json", help="Pfad zur Konfiguration (default: ehw_export.conf.json)")
    ap.add_argument("--mode", choices=["copy", "move", "symlink"], default="copy", help="Ablageart (default: copy)")
    ap.add_argument("--apply", action="store_true", help="Ohne dieses Flag nur Dry-Run.")
    ap.add_argument("--folder", action="append", help="Optional: nur diese(n) Folder verarbeiten; mehrfach möglich.")
    ap.add_argument("-q", "--quiet", action="store_true", help="Weniger Ausgabe.")
    args = ap.parse_args()

    conf = load_conf(Path(args.config))
    src_base = Path(conf["source_base_dir"]).resolve()
    dst_base = Path(conf["target_base_dir"]).resolve()
    folders = args.folder if args.folder else conf.get("folders", [])

    if not folders:
        print("No folders given (config.folders empty).")
        raise SystemExit(2)

    print(f"Source base: {src_base}")
    print(f"Target base: {dst_base}")
    print(f"Folders: {folders}")
    print(f"Mode: {args.mode} | Apply: {args.apply} | Verbose: {not args.quiet}")

    total_moved = total_skipped = total_errors = 0
    for name in folders:
        src_folder = src_base / name
        if not src_folder.exists():
            print(f"! Missing source folder: {src_folder}")
            continue
        m, s, e = process_folder(src_folder, dst_base, args.mode, args.apply, not args.quiet)
        total_moved += m
        total_skipped += s
        total_errors += e

    print(f"=== DONE === moved={total_moved} skipped={total_skipped} errors={total_errors}")

if __name__ == "__main__":
    main()
