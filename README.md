# Financial File Merger & Master Dashboard Builder

Streamlit app that consolidates NSE market-data CSV/Excel exports into one
formatted workbook, then automatically builds a `Master_Dashboard-8` sheet
joining every tab on `Symbol`.

## Files

- **`app.py`** — the Streamlit app. Run with:
  ```
  pip install -r requirements.txt
  streamlit run app.py
  ```
- **`MasterDashboardBuilder.bas`** — native VBA port of the same
  Master_Dashboard-8 build logic, for rebuilding it inside Excel directly
  (e.g. after manual edits). See `EXCEL_MACRO_SETUP.md` for setup.
- **`EXCEL_MACRO_SETUP.md`** — step-by-step instructions for importing the
  macro into a `.xlsm` workbook and wiring it to `Workbook_Open`.
- **`requirements.txt`** — Python dependencies.

## Optional login gate

`app.py` supports an optional login screen via Streamlit secrets. Create
`.streamlit/secrets.toml` (already git-ignored) with:
```toml
[auth]
username = "your_username"
password = "your_password"
```
If this file/section is absent, the app runs without a login prompt.
