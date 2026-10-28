# Fogwing Sync for Google Sheets

Google Apps Script that pulls weighment-scale records from the Fogwing Cloud
data-storage API into a Google Sheet, with a menu, a sidebar, a status check and
a sync log.

## Features
- **Manual fetch** by start/end date (sidebar or menu) and a one-click "Fetch Today Now"
- **Automatic end-of-day fetch** (daily time-driven trigger, around 11:30 PM IST)
- **De-duplication:** re-running a date range never adds the same record twice
- **Sync Log tab** recording every run (time, mode, range, fetched/added/skipped, status, error)
- **Status and connection test** from the menu or the sidebar
- Failure email to the sheet owner when the auto-fetch fails

## Files
| File | Purpose |
|---|---|
| `Code.gs` | All logic: menu, fetch, de-dupe, log, trigger |
| `Sidebar.html` | Fetch panel UI |
| `appsscript.json` | Project manifest (time zone Asia/Kolkata) |

## Setup
1. Create a Google Sheet, then open **Extensions > Apps Script**.
2. Add `Code.gs` and `Sidebar.html` (HTML file named exactly `Sidebar`). Optionally copy `appsscript.json` (enable *Show appsscript.json* in Project Settings).
3. Save and reload the Sheet. A **Fogwing Sync** menu appears; approve the permission prompts.
4. Run **Fogwing Sync > Set Credentials** and enter your API key, account ID and edge EUI.
   They are stored in Script Properties, never in the code or the repo.
5. Run **Test API Connection**, then **Turn ON Daily Auto-Fetch**.

## Data columns
Date, Material Type, Process Type, Shift Date, Shift Name, Shift Range, Source,
Time, Weight, Batch No, Vehicle No.

## Expected API response shape
```json
{
  "statusCode": 200,
  "message": "OK",
  "data": [
    {
      "payload": {
        "date": "01 Jan 2026",
        "time": "10:00:00 AM IST",
        "source": "Example source",
        "weight": 50.0,
        "batch_no": "01 Jan 2026 - 1st shift - 01",
        "shift_date": "01 Jan 2026",
        "shift_name": "1st shift",
        "vehicle_no": "XX00-0000",
        "shift_range": "08:00 - 05:00",
        "process_type": "Inward / Incoming",
        "material_type": "Example material"
      }
    }
  ]
}
```

## Security
- Never commit your API key, account ID or edge EUI. This repo reads them from Script Properties.
- Anyone with edit access to the Sheet's script can read Script Properties. Keep sharing limited.
- If a key is ever exposed, regenerate it in Fogwing and run **Set Credentials** again.
