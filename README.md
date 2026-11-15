# 406 Beacon Registration Script

Registers 406 MHz beacons on <https://register-406-beacons.service.gov.uk> from an Excel sheet, one row per beacon, asking before each one.

## Files

| File | Purpose |
|---|---|
| `register_from_excel.py` | The script: reads the sheet and registers each row |
| `beacons_template.xlsx` | Template with the required columns |
| `results.csv` | Created on first run; log of every outcome |

## Setup

```powershell
pip install requests openpyxl
```

## Fill in the sheet

One beacon per row, with these columns (all required):

`manufacturer`, `model`, `hexId`, `serialNumber`, `environment`, `activityDescription`, `moreDetails`, `ownerName`, `ownerTelephone`, `ownerEmail`, `addressLine1`, `townOrCity`, `postcode`, `contactName`, `contactTelephone`

Keep phone-number cells as text so they keep their leading 0.

Every beacon is registered as **`environment` → Commercial → Other**, with `activityDescription` as the description. It has one use per beacon and a UK owner address.

`environment` accepts `MARITIME` or `AVIATION`. Rows with `LAND` (or anything else) are skipped with a message, and nothing is sent to the site.

**LAND is out of scope for this first version.** The site requires a location for a land beacon, which this sheet doesn't collect. Register land beacons by hand on the site for now.

## Get your cookie

The script signs in using your browser's session cookie:

1. Sign in at <https://register-406-beacons.service.gov.uk>.
2. Press **F12** and open the **Network** tab.
3. Refresh the page (**F5**).
4. Click the first request in the list.
5. Under **Request Headers**, right-click **cookie** and choose **Copy value**.

Use the Network tab, not `document.cookie` in the Console: the sign-in cookies are hidden from it. The cookie expires when you sign out or your session times out. The script asks for a new one if that happens.

## Run

```powershell
python register_from_excel.py beacons.xlsx
```

When asked, paste the cookie and press Enter. Or press Enter on an empty line to read it from the clipboard, or from a `cookie.txt` file in this folder. Optional settings:

- `$env:BEACON_COOKIE = '...'` skips the cookie prompt.
- `$env:BEACON_DEBUG = "1"` prints every request's status, redirects and cookie names.
- `--sheet NAME` selects the sheet to read.
- `--results FILE` changes the log file.

For each row it shows the values and asks **[y]es / [s]kip / [q]uit**.

## What the script does

1. Checks you're signed in (`GET /api/auth/session`). If not, it asks for a fresh cookie.
2. Reads the sheet. It skips rows with blank fields, rows with an unsupported `environment`, and hex IDs already logged as `registered` in `results.csv`.
3. For each row you confirm:
   1. Clears any saved draft (`GET /api/registration/clear-and-check-beacon-details`).
   2. Submits each form page in order, following each redirect like a browser:
      1. `check-beacon-details`: manufacturer, model, hex ID
      2. `beacon-information`: serial number
      3. `beacon-use?useId=0`: `environment`
      4. `purpose?useId=0`: Commercial
      5. `activity?useId=0`: Other plus the description
      6. `about-the-vessel?useId=0` (MARITIME) or `about-the-aircraft?useId=0` (AVIATION): left blank
      7. `vessel-communications?useId=0` (MARITIME) or `aircraft-communications?useId=0` (AVIATION): left blank
      8. `more-details?useId=0`
      9. `about-beacon-owner`: name, telephone, email
      10. `beacon-owner-address`: United Kingdom
      11. `beacon-owner-address/united-kingdom`: line 1, town, postcode
      12. `emergency-contact`: contact 1
   3. A page that returns 303 has been accepted. A page that returns 200 has been rejected: the script prints the site's error, nothing is submitted, and it moves on to the next row.
   4. Submits with a single `GET /register-a-beacon/application-complete` and reads the reference number from the page.
4. Logs the outcome to `results.csv`: time, row, hex ID, status, reference, detail.

## Safety

- The submit is **never retried**. If it fails or times out, the script stops the run, because the beacon may already be registered. Check your account's beacon list before running again.
- Re-running the same sheet is safe: rows already logged as registered in `results.csv` are skipped.

### Duplicate hex IDs are not prevented

The HM Coastguard site accepts a hex ID that is already registered and creates a second registration. This is a known limitation of the site. The script only skips hex IDs that are logged as `registered` in its own `results.csv`. It doesn't check your account, so it will register a duplicate if the beacon was:

- registered by hand on the site, or by someone else,
- registered from another computer or folder, or before `results.csv` was deleted or moved.

Before a run, check your account's beacon list and remove any rows that are already registered there.
- The cookie gives access to your account. Don't share it, and delete `cookie.txt` after use.
