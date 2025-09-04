# valorant-analytical-reports
Python application that automatically builds a multi-sheet Google Spreadsheet report for a given Valorant team and a selected set of recent matches. It pulls data from your internal data functions (via functions.py), formats it, uploads generated images (early positioning & sniper kills) to Drive, and embeds everything into a styled Google Sheet.

## What it does

- Interactive CLI: choose a **team tag** (e.g., `TH`) and how many **most-recent matches** to include.
- Fetches match data via your `functions.py` helpers.
- Creates a **Google Spreadsheet**:
  - **Overall** sheet with: results table, DEF/ATK side winrates, and a **“Performance by Map”** summary.
  - One **sheet per map** with:
    - Agent compositions & winrates.
    - Post-plant and pistol post-plant performance (ATK/DEF).
    - Early team positioning (10s/20s/30s) with embedded images.
    - Sniper kills (ATK/DEF) with embedded images.
- Auto-styles headers/cells and **embeds the team logo**.
- Shares the spreadsheet with your email and prints the final URL.

## Repo structure (key files)
```text
.
├─ README.md
├─ main.py                             # Entry-point script that builds the spreadsheet
├─ functions.py                        # Data layer helpers 
├─ plots/                              # Auto-generated images before upload to Drive
└─ .gitignore                          # Should ignore credentials & caches
