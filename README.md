# Fees Processing Application

This repository contains a Flask web application for processing Excel workbooks. The project helps automate fee calculations by uploading a workbook and receiving a processed file in return.

## Structure

- `app/` – Minimal Flask application entry point used by `run.py`.
- `fees/` – Self-contained package with the same functionality, including a blueprint and tests.
- `templates/` – HTML templates for the root application.
- `requirements.txt` – Python dependencies.
- `run.py` – Starts the root application.

See `fees/README.md` for detailed usage instructions and development notes.
