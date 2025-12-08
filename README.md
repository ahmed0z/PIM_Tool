# PIM_Tool

PIM Format Automation Tool - A Streamlit web application for processing PIM Issue Reports.

## Features

- **Main Page**: Upload PIM file and Part Data file for processing
- **Settings Page**: Manage preset database (upload/update from Excel)
- Preset database is stored in the repo to avoid re-uploading

## Installation

```bash
pip install -r requirements.txt
```

## Usage

1. **First Run**: Go to Settings page and upload your preset Excel file
2. **Processing**: On the main page, upload your PIM and Part Data files
3. Click "Run Process" and download the results

### Run the app

```bash
streamlit run app.py
```

## Files

- `app.py` - Main Streamlit application
- `preset_db.pkl` - Preset database (created after first upload)
- `requirements.txt` - Python dependencies
