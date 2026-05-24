# FAI Dimension Numberer

A desktop tool for annotating engineering drawings (PDFs) with numbered callout circles for First Article Inspection (FAI) reports.

## Features
- Automatically detects dimensions from PDF drawings using pattern matching
- Overlays numbered callout circles, ordered left-to-right / top-to-bottom
- Detects general tolerances from the drawing's title block
- Toggle, move, delete, and manually add callouts
- Undo (Ctrl+Z) for moves and toggles
- Re-open previously annotated drawings and resume editing (session saved alongside PDF)
- Export annotated PDF and filled Excel measurement record

## Running from source

### Requirements
- Python 3.11+
- Dependencies listed in `requirements.txt`

```
pip install -r requirements.txt
python fai_dimension_numberer.py
```

## Building the .exe

Run the included build script (Windows):

```
build.bat
```

Output: `dist\FAI_Dimension_Numberer.exe`

Copy `MSI-VSA_Template_Dim_Check.xlsx` into the same `dist\` folder before distributing.

## Files
| File | Purpose |
|---|---|
| `fai_dimension_numberer.py` | Main application (UI, export, session management) |
| `dimension_detector.py` | Regex-based dimension and tolerance detection |
| `MSI-VSA_Template_Dim_Check.xlsx` | Measurement record template (required at runtime) |
| `requirements.txt` | Python dependencies |
| `build.bat` | PyInstaller build script |
