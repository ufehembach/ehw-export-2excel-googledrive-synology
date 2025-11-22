

# EHW+ JSON → Excel Exporter  
Konvertiert EHW+ (EffizienzHausWächter) Google-Drive/JSON-Exporte in saubere, analysierbare Excel-Dateien.

## ✨ Features
- Liest komplette EHW+ JSON-Dumps (inkl. Rooms, Counters, Entries)
- Bildexport: Kopiert oder symlinkt Zählerfotos aus dem `.hidden` Google-Drive Ordner
- **Delta-Berechnung** (inkl. Reset-Erkennung)
  - PrevValue  
  - PrevDate  
  - Delta  
  - Days  
  - DeltaPerDay
- Monatliche & jährliche Aggregation
- Automatische Excel-Tabellenformatierung (openpyxl)
- Kombinierter Export mehrerer Ordner in `ehw+.xlsx`
- Unterstützung für virtuelle Zähler (Vorbereitung vorhanden)
- Vollständige Struktur kompatibel mit Pivot & PowerBI

## 📁 Ordnerstruktur
```
ehw_export/
├── ehw_export.py
├── ehw_transform.py
├── ehw_fix_Images.py
├── ehw_export_augment.py
├── ehw_export.conf.json
├── VERSION
└── myEHW+GoogleDrive/
```

## 🔧 Konfiguration
Die Datei **ehw_export.conf.json** definiert:
- Quellverzeichnis (Google-Drive Sync)
- Zielverzeichnis (Synology/Ordner)
- Welche Ordner exportiert werden sollen

Beispiel:
```json
{
  "source_base_dir": "/Volumes/GoogleDrive/EHW+",
  "target_base_dir": "/volume1/ehw_export",
  "folders": ["DBMP", "H1", "H3"]
}
```

## ▶️ Nutzung
```
./ehw_export.py
```

Dies erzeugt:
- `##DBMP-YYYYmmdd_HHMMSS.xlsx`
- `DBMP.xlsx` (always latest)
- `ehw+.xlsx` (kombiniert alle Ordner)

## 🧮 Excel-Sheets
### 1. **Zählerdaten**  
Alle Rohdaten + Delta-Informationen  
→ Tabelle: `tblEHW`

### 2. **Zählerdaten_Monat**  
Monatliche Aggregation (inkl. Delta)  
→ Tabelle: `tblehwMonat`

### 3. **Zählerdaten_Jahr**  
Jährliche Aggregation (inkl. Delta)  
→ Tabelle: `tblehwJahr`

## 🔥 Delta-Berechnung
Delta wird automatisch aus `Value_Num` berechnet:
- Wenn der neue Wert **kleiner** ist als der alte → *Reset*  
  → Delta = neuer Wert  
  → PrevValue = None  
- Sonst  
  → Delta = Value – PrevValue

### Verbrauch pro Tag:
```
DeltaPerDay = Delta / Days
```

## 🧩 Virtuelle Zähler
Struktur in JSON:
```
"counterType": "VIRTUAL",
"virtualCounterData": {
  "masterCounterUuid": "...",
  "counterUuidsToBeAdded": [...],
  "counterUuidsToBeSubtracted": [...]
}
```
→ Vorbereitung im Code vorhanden  
→ Implementierung folgt (additive/subtraktive Berechnung)

## 🛠 Versionierung
Die Datei `VERSION` enthält die aktuelle Versionsnummer.
Diese wird im Excel-Header angezeigt.

## 📌 TODO / Roadmap
- [ ] Virtuelle Zähler vollständig berechnen
- [ ] Automatische VERSION-Erhöhung (optional)
- [ ] Performance-Optimierung für große Exporte
- [ ] Markdown-basierte Release Notes
- [ ] GitHub Actions CI

## © Lizenz
Persönliches Projekt von **ufehembach**  
Keine Gewährleistung, Nutzung auf eigene Verantwortung.