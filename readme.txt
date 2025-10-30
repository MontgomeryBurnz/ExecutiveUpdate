# Program Status (Exec Packet)

Interactive Streamlit app for weekly status updates. Upload the prior Excel packet, collaborate on edits directly in the browser, review quick quality checks, and export an updated packet or executive scorecard view.

## Features
- Upload any Excel workbook (`.xlsx`); every sheet becomes an editable Streamlit tab.
- RAG-select helpers keep status values consistent for the roll-up.
- Executive Scorecard summarizes overall status counts and upcoming/overdue milestones.
- Sidebar quality checks surface nulls in key columns and date parsing issues across all sheets.
- Export the refreshed Excel packet or the primary sheet as CSV on demand.
- Optional Snowflake sync scaffold (commented) shows how to push the curated data with `write_pandas`.

## Quick Start
1. Create a virtual environment and install dependencies:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
2. Launch the Streamlit app:
   ```bash
   streamlit run WeeklyUpdate.py
   ```
3. Upload last week’s Excel packet to begin editing. If none is available you can create a blank starter sheet from the UI.

## Optional Snowflake Sync
The Snowflake helper in `WeeklyUpdate.py` is commented out by default. To enable it:
1. Add a `[snowflake]` section to `.streamlit/secrets.toml` with credentials (account, user, password, warehouse, database, schema).
2. Install the optional dependency:
   ```bash
   pip install snowflake-connector-python
   ```
3. Un-comment the `write_to_snowflake` function and wire it to a button or trigger of your choice.

## Testing
Run a basic syntax check before deploying:
```bash
python3 -m compileall WeeklyUpdate.py
```
For richer validation, add unit tests around the helper functions or run Streamlit locally and exercise the workflow end-to-end.
