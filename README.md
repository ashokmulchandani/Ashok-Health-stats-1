# Ashok Health Stats — Daily Entry Form

A small web form to punch daily data (weight, body fat, sleep, steps, HIIT, water, comment) straight into the workout workbook in the OneDrive folder. OneDrive syncs automatically afterwards.

## How to use (PC)
1. Double-click **`start_ashok_form.bat`**
2. The form opens at **http://localhost:8765** — type the day's numbers and press **Save to workbook**
3. The connector writes ONLY the fields you typed into `Daily Metrics Metabolism` (your rules: never touches anything else, safe-save with backup behaviour), then OneDrive syncs

## How to use (phone — same Wi-Fi)
1. Keep `start_ashok_form.bat` running on the PC (PC must be ON; locked is fine)
2. In the black connector window it prints a link like `http://192.168.x.x:8765`
3. Open that link on your phone → same form, mobile-friendly → Save → workbook updates on the PC → OneDrive syncs

> First time on the phone, Windows may pop up a firewall prompt on the PC — click **Allow** (private networks).

## Notes
- To change the workbook path, create `config.json` next to the scripts:
  `{ "workbook": "C:\\full\\path\\to\\file.xlsx" }`
- Do **not** have the workbook open in Excel while saving — the connector will refuse and ask you to close it.
- The GitHub Pages copy of this form (repo page) is for viewing only — saving works only from the connector's own page.

## Files
- `index.html` — the form (mobile friendly)
- `connector.py` — local server + workbook writer (Python 3 only, no extra packages)
- `start_ashok_form.bat` — double-click launcher
