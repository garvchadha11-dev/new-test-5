# FTA Excise Portal Scraper v5

Automated tool for downloading declaration reports from the UAE FTA Excise Tax portal. Built for Andersen Consulting internal use.

**v5 changes from v4:**
- Added EX200 (Excise Tax Return) and Excise 311 (Direct Refunds) declaration types
- Filter fallback improved: when Approved status returns no records, automatically retries with Warehouse Keeper status before giving up (previously it would give up immediately)

---

## Requirements

- Windows 10 / 11
- Python 3.10 or newer — download from [python.org](https://www.python.org/downloads/)
  - During install, check **"Add Python to PATH"**
- Microsoft Edge (already installed on most Windows machines)

---

## How to Run

1. Double-click **`run.bat`**
2. It will automatically install the required packages (`playwright`, `openpyxl`) and launch the app.

---

## How to Use

### Step 1 — Open Browser & Login
Click **"1. Open Browser & Login"**. A Microsoft Edge window will open and navigate to the FTA portal.  
Log in manually with your FTA credentials. Do not close that Edge window.

### Step 2 — Configure Settings

| Setting | What it does |
|---|---|
| **Date Range** | Select the From and To month/year. The scraper will process every month in that range. |
| **Declaration Types** | Tick the declaration forms you want to export. Use "Select All" / "Clear All" for bulk selection. |
| **Save Folder** | Choose where downloaded files will be saved. Defaults to your Downloads folder. |

### Step 3 — Start Scraping
Click **"2. Start Scraping"**. The tool will:
- Navigate each selected declaration panel
- Apply the date filter for each month in your range
- Download rows one by one into organised subfolders
- Combine all downloaded files into a single merged Excel file at the end

You can click **Stop** at any time — it will finish the current row cleanly before stopping.

---

## Output Structure

```
Downloads/
└── 2025-01-15_14-30/          ← timestamped run folder
    ├── EX200/
    │   ├── January 2025/
    │   │   └── *.xlsx
    │   └── February 2025/
    ├── EX201_ML/
    │   └── ...
    └── Combined_Report.xlsx   ← all data merged
```

---

## Security Notice

> **JavaScript Injection Risk**
>
> This tool passes user-selected values (month names, years) directly into JavaScript strings that are executed inside the browser via Playwright. The current values are safe because they come from a fixed dropdown list. However, if this tool is ever modified to accept free-text input, those values **must be escaped before insertion into JS strings** — otherwise a malicious or malformed input could break the script or execute unintended code in the browser context.
>
> Similarly, file paths entered in the Save Folder field are passed to `os.makedirs` and used in `subprocess` calls. Avoid paths with unusual characters or names crafted to look like shell commands if running in an environment where the folder input could come from an untrusted source.
>
> **In short:** keep inputs coming from the fixed UI controls and do not expose this tool to untrusted input without adding proper sanitisation first.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "Python is not installed" | Reinstall Python and check "Add to PATH" |
| Browser opens but scraper can't connect | Make sure no other Edge window is using port 9222 (e.g. close VS Code browser tools) |
| "No data after filtering" warning | That month has no records on the portal — this is normal |
| App freezes / Start button stays greyed | Click Stop, close the app, and restart via `run.bat` |

---

*Report issues to Garv.*
