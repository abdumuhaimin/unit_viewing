# Unit Viewing Dashboard

This project turns `property_listings_MM.xlsx` into a Streamlit dashboard for reviewing listings and saving your unit-viewing checklist.

## Live App

https://mxmunitviewing.streamlit.app/

The GitHub repo does not store the source workbook. On deployed environments, upload `property_listings_MM.xlsx` when the app prompts for it.

## What it does

- Loads the workbook into one dashboard view.
- Lets you filter by rent, MRT station, MRT distance, property type, furnishing, co-living, and review status.
- Adds a persistent checklist for each listing covering:
  - Utilities + air-con inclusion
  - Water heater
  - Washing machine
  - Fridge
  - Bedroom ceiling fan
  - Living room ceiling fan
  - Floor level
  - Sunrise or sunset exposure
  - Building crowd profile
  - Common kitchen setup for co-living units
- Saves review notes to `unit_viewing_reviews.csv`.

## Run locally

1. Install dependencies:

   ```powershell
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

2. Start the dashboard:

   ```powershell
   .\.venv\Scripts\python.exe -m streamlit run app.py
   ```

3. Open the local Streamlit URL shown in the terminal.

## Files

- `app.py` - dashboard app
- `property_listings_MM.xlsx` - source workbook kept locally and intentionally excluded from git
- `unit_viewing_reviews.csv` - saved checklist data, created after your first save