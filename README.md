# PowerPoint self-QC application

A Windows-oriented Flask application for checking PowerPoint presentation quality and generating QC reports.

## Included checks

- Text and formatting rules
- Animation-related checks and slide chunking
- Slide notes validation
- US to UK spelling checks
- QC summary generation

## Requirements

- Python on Windows
- Microsoft PowerPoint for Windows-specific automation (when required by the win32 modules)
- Dependencies listed in `requirements.txt`

Some dependencies, particularly `aspose.slides`, `torch`, `sentence-transformers`, and `pywin32`, can have platform, licensing, or installation requirements. The `tk` entry in requirements may not install from pip; Tkinter is normally installed with suitable Python builds.

## Run locally

From this folder in PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

Alternatively, run `run_app.bat` on Windows. It creates a `venv` folder and installs the dependencies. Open the localhost address printed by Flask in your browser.

## Project structure

- `app.py` — Flask app entry point
- `templates/` — HTML page
- `*_checker.py`, `*_validator.py`, `qc_points_generator.py` — QC modules
- `us_to_uk_dictionary.csv` — spelling data
- `run_app.bat` — Windows launcher

## Notes

Do not commit private presentation files, reports, uploaded documents, environment files, or credentials. Check the bundled deployment guide and sample spreadsheet for information you do not intend to share before publishing the repository.

This project is provided as uploaded; end-to-end functional testing has not been performed.
