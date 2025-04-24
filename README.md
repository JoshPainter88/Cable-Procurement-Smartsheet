# Cable-Procurement-Smartsheet

Python helpers and API scripts that sync our Smartsheet cable-procurement tracker
with a structured database (plus future Power BI loads).

## What’s here
* `scripts/` batch scripts that call the Smartsheet REST API  
* `data-models/` ER-diagrams, CSV templates for each sheet  
* `docs/` setup notes and FAQ

## Quick start
1. Create a Smartsheet API token and store it in an environment variable  
   `SMARTSHEET_TOKEN` — see `docs/auth.md`  
2. Install requirements: `pip install -r requirements.txt`  
3. Run `python scripts/pull_inventory.py`

**WIP** – first goal is to pull Change Orders automatically.
