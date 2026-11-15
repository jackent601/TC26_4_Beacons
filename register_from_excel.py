#!/usr/bin/env python3
"""
Register beacons listed in an Excel sheet, one row per beacon, asking before each one.

Usage:
    python3 register_from_excel.py beacons.xlsx [--sheet Sheet1] [--results results.csv]

It asks you to paste your browser's Cookie header at the prompt (pressing Enter
on an empty line reads the clipboard, then cookie.txt next to this script instead),
unless the BEACON_COOKIE environment variable is already set. If the session expires
mid-run it asks again rather than stopping.

For each row it prints what it will register and asks:
    y = register it    s = skip it    q = stop
Every outcome is appended to the results CSV. Hex IDs already recorded there as
registered are skipped automatically, so re-running the same sheet is safe.

Needs: pip install requests openpyxl
"""
import argparse
import html
import json
import re
import csv
import datetime as dt
import os
import sys

import openpyxl
import requests


# --------------------------------------------------------------- site requests
BASE = os.environ.get("BEACON_BASE_URL", "https://register-406-beacons.service.gov.uk").rstrip("/")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36")


class StepFailed(Exception):
    pass


def set_cookies(s: requests.Session, cookie: str):
    """Load a pasted Cookie header into the session's cookie jar (not a fixed header),
    so cookies the site sets during the run (e.g. a new draft id) replace the old ones."""
    host = requests.compat.urlparse(BASE).hostname
    s.cookies.clear()
    for part in cookie.split(";"):
        if "=" in part:
            name, value = part.strip().split("=", 1)
            if name:
                s.cookies.set(name, value, domain=host, path="/")


def make_session(cookie: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Origin": BASE,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-GB,en;q=0.9",
    })
    set_cookies(s, cookie)
    if os.environ.get("BEACON_DEBUG"):
        def log(r, *a, **k):
            for x in list(r.history) + [r]:
                names = [c.split("=", 1)[0] for c in x.raw.headers.getlist("Set-Cookie")] \
                    if hasattr(x.raw, "headers") and hasattr(x.raw.headers, "getlist") else []
                print(f"    [debug] {x.request.method} {x.status_code} "
                      f"{x.url.replace(BASE, '')}"
                      + (f"  -> {x.headers.get('Location')}" if x.is_redirect else "")
                      + (f"  set-cookie: {', '.join(names)}" if names else ""))
        s.hooks["response"].append(log)
    return s


def page_errors(body: str) -> list[str]:
    """Pull messages out of the GOV.UK error summary on a rejected page."""
    m = re.search(r'class="govuk-error-summary.*?</div>\s*</div>', body, re.S)
    chunk = m.group(0) if m else body
    items = re.findall(r"<li[^>]*>(.*?)</li>", chunk, re.S)
    msgs = [html.unescape(re.sub(r"<[^>]+>", "", i)).strip() for i in items]
    return [x for x in msgs if x] or ["(page returned 200 but no error summary found)"]


def get(s, path, **kw):
    r = s.get(BASE + path, timeout=30, **kw)
    r.raise_for_status()
    return r


def post_step(s, path, fields: dict):
    """POST one form page. Success = 303 redirect; 200 = validation errors."""
    url = BASE + path
    get(s, path)  # load the page first, as the browser does
    r = s.post(url, data=fields, allow_redirects=False, timeout=30,
               headers={"Referer": url,
                        "Content-Type": "application/x-www-form-urlencoded"})
    if r.status_code in (302, 303):
        loc = r.headers.get("Location", "")
        # Follow the redirect like a browser: some of these GETs change server state
        # (e.g. GET /register-a-beacon/beacon-use creates use 0, then redirects to ?useId=0).
        if loc:
            nxt = s.get(requests.compat.urljoin(url, loc), timeout=30)
            nxt.raise_for_status()
        print(f"  ok   {path}")
        return loc
    if r.status_code == 200:
        raise StepFailed(f"{path} rejected:\n    - " + "\n    - ".join(page_errors(r.text)))
    raise StepFailed(f"{path}: unexpected HTTP {r.status_code}")


def s_(d, k):
    """Value as a string; missing/None -> ''."""
    v = d.get(k)
    return "" if v is None else str(v)


def build_steps(cfg: dict):
    b = cfg["beacon"]
    u = cfg["use"]
    v = u.get("vessel", {})
    c = u.get("communications", {})
    o = cfg["owner"]
    a = o["address"]
    contacts = cfg["emergencyContacts"]
    if not 1 <= len(contacts) <= 3:
        raise ValueError("emergencyContacts must have 1 to 3 entries")
    env = u.get("environment", "MARITIME").strip().upper()
    if env not in SUPPORTED_ENVIRONMENTS:
        raise ValueError(f"environment {env!r} isn't supported yet "
                         f"(supported: {', '.join(SUPPORTED_ENVIRONMENTS)})")

    # Default use is what was recorded: commercial, activity OTHER + a description.
    u.setdefault("purpose", "COMMERCIAL")
    u.setdefault("activity", "OTHER")
    if u["activity"] == "OTHER" and not s_(u, "otherActivityText").strip():
        raise ValueError("activity is OTHER, so use.otherActivityText (the description) is required")

    uid = "0"
    q = f"?useId={uid}"

    check = {"manufacturer": s_(b, "manufacturer"), "model": s_(b, "model"),
             "hexId": s_(b, "hexId").upper(), "useId": uid}
    if b.get("isSecondGeneration"):
        check["isSecondGeneration"] = "true"

    comms = {}
    # each comms entry: checkbox name -> its text input
    for box, inp in [("fixedVhfRadio", "fixedVhfRadioInput"),
                     ("portableVhfRadio", "portableVhfRadioInput"),
                     ("satelliteTelephone", "satelliteTelephoneInput"),
                     ("otherCommunication", "otherCommunicationInput")]:
        if c.get(inp):
            comms[box] = "true"
    if c.get("fixedVhfRadioInput") or c.get("portableVhfRadioInput"):
        comms["vhfRadio"] = "true"
    if c.get("mobileTelephoneInput1") or c.get("mobileTelephoneInput2"):
        comms["mobileTelephone"] = "true"
    for k in ["fixedVhfRadioInput", "portableVhfRadioInput", "satelliteTelephoneInput",
              "mobileTelephoneInput1", "mobileTelephoneInput2",
              "otherCommunicationInput", "callSign"]:
        comms[k] = s_(c, k)
    comms["useId"] = uid

    ec = {}
    for i in range(1, 4):
        p = contacts[i - 1] if i <= len(contacts) else {}
        ec[f"emergencyContact{i}FullName"] = s_(p, "fullName")
        ec[f"emergencyContact{i}TelephoneNumber"] = s_(p, "telephone")
        ec[f"emergencyContact{i}AlternativeTelephoneNumber"] = s_(p, "alternativeTelephone")

    return [
        ("/register-a-beacon/check-beacon-details", check),
        ("/register-a-beacon/beacon-information", {
            "manufacturerSerialNumber": s_(b, "manufacturerSerialNumber"),
            "chkCode": s_(b, "chkCode"), "csta": s_(b, "csta"),
            "batteryExpiryDateMonth": s_(b, "batteryExpiryDateMonth"),
            "batteryExpiryDateYear": s_(b, "batteryExpiryDateYear"),
            "lastServicedDateMonth": s_(b, "lastServicedDateMonth"),
            "lastServicedDateYear": s_(b, "lastServicedDateYear"),
            "useId": uid}),
        ("/register-a-beacon/beacon-use" + q, {"environment": env, "useId": uid}),
        ("/register-a-beacon/purpose" + q, {"purpose": s_(u, "purpose"), "useId": uid}),
        ("/register-a-beacon/activity" + q, {
            "activity": s_(u, "activity"), "otherActivityText": s_(u, "otherActivityText"),
            "environment": env, "useId": uid}),
        ("/register-a-beacon/about-the-vessel" + q, {
            k: s_(v, k) for k in ["maxCapacity", "vesselName", "beaconLocation",
                                  "portLetterNumber", "homeport", "areaOfOperation",
                                  "imoNumber", "ssrNumber", "rssNumber",
                                  "officialNumber", "rigPlatformLocation"]} | {"useId": uid}),
        ("/register-a-beacon/vessel-communications" + q, comms),
        ("/register-a-beacon/more-details" + q, {"moreDetails": s_(u, "moreDetails"), "useId": uid}),
        # additional-beacon-use has no form; "Continue" is a plain link to about-beacon-owner
        ("/register-a-beacon/about-beacon-owner", {
            "ownerFullName": s_(o, "fullName"), "ownerTelephoneNumber": s_(o, "telephone"),
            "ownerAlternativeTelephoneNumber": s_(o, "alternativeTelephone"),
            "ownerEmail": s_(o, "email"), "useId": uid}),
        ("/register-a-beacon/beacon-owner-address", {"beaconOwnerLocation": "unitedKingdom", "useId": uid}),
        ("/register-a-beacon/beacon-owner-address/united-kingdom", {
            "ownerAddressLine1": s_(a, "line1"), "ownerAddressLine2": s_(a, "line2"),
            "ownerTownOrCity": s_(a, "townOrCity"), "ownerCounty": s_(a, "county"),
            "ownerPostcode": s_(a, "postcode")}),
        ("/register-a-beacon/emergency-contact", ec),
    ]


def next_data(body: str) -> dict:
    m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', body, re.S)
    if not m:
        raise StepFailed("no __NEXT_DATA__ on application-complete page")
    return json.loads(m.group(1))



# --------------------------------------------------------------- sheet handling
# Site values: MARITIME, AVIATION, LAND. Only MARITIME's pages have been recorded;
# AVIATION and LAND lead to different pages and need a HAR of that flow to add.
SUPPORTED_ENVIRONMENTS = ("MARITIME",)

# column header -> description (all required). Headers are matched case-insensitively.
COLUMNS = {
    "manufacturer": "Beacon manufacturer",
    "model": "Beacon model",
    "hexId": "Beacon HEX ID / UIN",
    "serialNumber": "Manufacturer serial number",
    "environment": "Environment (MARITIME)",
    "activityDescription": "Description of the use (activity = Other)",
    "moreDetails": "More details that could help in a search",
    "ownerName": "Owner full name",
    "ownerTelephone": "Owner telephone",
    "ownerEmail": "Owner email",
    "addressLine1": "Owner address line 1",
    "townOrCity": "Owner town or city",
    "postcode": "Owner postcode",
    "contactName": "Emergency contact full name",
    "contactTelephone": "Emergency contact telephone",
}


def cell_str(v) -> str:
    """Excel cell -> clean string. Keeps leading zeros only if the cell is text."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def read_rows(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = [cell_str(h) for h in next(rows)]
    lookup = {h.lower(): i for i, h in enumerate(header) if h}
    missing = [c for c in COLUMNS if c.lower() not in lookup]
    if missing:
        sys.exit(f"Sheet is missing column(s): {', '.join(missing)}\nFound: {header}")
    out = []
    for n, row in enumerate(rows, start=2):
        rec = {c: cell_str(row[lookup[c.lower()]]) if lookup[c.lower()] < len(row) else ""
               for c in COLUMNS}
        if any(rec.values()):
            out.append((n, rec))
    return out


def to_config(r: dict) -> dict:
    return {
        "beacon": {"manufacturer": r["manufacturer"], "model": r["model"],
                   "hexId": r["hexId"], "manufacturerSerialNumber": r["serialNumber"]},
        "use": {"environment": r["environment"].upper(),
                "otherActivityText": r["activityDescription"], "moreDetails": r["moreDetails"]},
        "owner": {"fullName": r["ownerName"], "telephone": r["ownerTelephone"],
                  "email": r["ownerEmail"],
                  "address": {"line1": r["addressLine1"], "townOrCity": r["townOrCity"],
                              "postcode": r["postcode"]}},
        "emergencyContacts": [{"fullName": r["contactName"], "telephone": r["contactTelephone"]}],
    }


def already_registered(results_path) -> dict:
    done = {}
    if os.path.exists(results_path):
        with open(results_path, newline="") as f:
            for row in csv.DictReader(f):
                if row.get("status") == "registered":
                    done[row["hexId"].upper()] = row.get("reference", "")
    return done


def log_result(results_path, row_no, hex_id, status, reference="", detail=""):
    new = not os.path.exists(results_path)
    with open(results_path, "a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["time", "row", "hexId", "status", "reference", "detail"])
        w.writerow([dt.datetime.now().isoformat(timespec="seconds"), row_no,
                    hex_id.upper(), status, reference, detail])


def clean_cookie(v: str) -> str:
    v = v.strip().strip("'\"").strip()
    if v.lower().startswith("cookie:"):
        v = v[7:].strip()
    return v


COOKIE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cookie.txt")


def read_clipboard() -> str:
    try:
        import tkinter
        root = tkinter.Tk()
        root.withdraw()
        try:
            return root.clipboard_get()
        finally:
            root.destroy()
    except Exception:
        pass
    try:  # Windows fallback
        import subprocess
        return subprocess.run(["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"],
                              capture_output=True, text=True, timeout=15).stdout
    except Exception:
        return ""


def looks_like_cookie(v: str) -> bool:
    return "=" in v and "\n" not in v.strip() and len(v) > 20


def ask_cookie() -> str:
    """Get the Cookie header: pasted at the prompt, else the clipboard, else cookie.txt."""
    print("\nNeed the Cookie header from your signed-in browser:")
    print("  F12 > Network > refresh the page > click the first request >")
    print("  Request Headers > right-click 'cookie' > Copy value")
    while True:
        try:
            v = clean_cookie(input("Paste it and press Enter (or just Enter to read the clipboard): "))
        except (KeyboardInterrupt, EOFError):
            sys.exit("\nStopped.")
        src = "terminal"
        if not v:
            v, src = clean_cookie(read_clipboard()), "clipboard"
        if not looks_like_cookie(v) and os.path.exists(COOKIE_FILE):
            v, src = clean_cookie(open(COOKIE_FILE, encoding="utf-8-sig").read()), "cookie.txt"
        if looks_like_cookie(v):
            print(f"  got {len(v)} characters from the {src}")
            return v
        print(f"  That doesn't look like a cookie (from the {src}; should be one long line of name=value; pairs).")
        print(f"  Try again, or save it in {COOKIE_FILE} and press Enter on an empty line.")


def show(row_no, r):
    print(f"\n──────── Row {row_no} ────────")
    for c, label in COLUMNS.items():
        print(f"  {label:<45} {r[c] or '(blank!)'}")


def fill_and_submit(s, cfg):
    """Returns (status, reference, detail). Never retries the submit."""
    steps = build_steps(cfg)
    get(s, "/api/registration/clear-and-check-beacon-details")
    for path, fields in steps:
        post_step(s, path, fields)
    try:
        r = s.get(BASE + "/register-a-beacon/application-complete", timeout=120,
                  headers={"Referer": BASE + "/register-a-beacon/check-your-answers"})
    except requests.RequestException as e:
        return ("UNKNOWN", "", f"submit request failed ({e}); check your account before retrying")
    if r.status_code != 200:
        return ("UNKNOWN", "", f"submit HTTP {r.status_code}; check your account before retrying")
    props = next_data(r.text).get("props", {}).get("pageProps", {})
    if not props.get("registrationSuccess"):
        return ("failed", "", f"site reported failure: {props}")
    detail = "" if props.get("confirmationEmailSuccess") else "confirmation email NOT sent"
    return ("registered", props.get("reference", ""), detail)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--sheet", help="sheet name (default: first sheet)")
    ap.add_argument("--results", default="results.csv")
    args = ap.parse_args()

    s = make_session(clean_cookie(os.environ.get("BEACON_COOKIE", "")) or ask_cookie())

    def check_signed_in():
        """Check the session; if it's missing or expired, ask for a fresh cookie."""
        while True:
            sess = get(s, "/api/auth/session", headers={"Accept": "application/json"}).json()
            if sess.get("user"):
                return sess
            print("\nNot signed in: the cookie is missing, incomplete or expired.")
            set_cookies(s, ask_cookie())

    sess = check_signed_in()
    print(f"Signed in as {sess['user'].get('email') or sess['user'].get('name')}")

    rows = read_rows(args.xlsx, args.sheet)
    done = already_registered(args.results)
    print(f"{len(rows)} row(s) in sheet; {len(done)} hex ID(s) already registered in {args.results}")

    counts = {"registered": 0, "skipped": 0, "failed": 0}
    for row_no, r in rows:
        hex_id = r["hexId"].upper()
        if hex_id in done:
            print(f"\nRow {row_no}: {hex_id} already registered (ref {done[hex_id]}); skipping.")
            counts["skipped"] += 1
            continue

        show(row_no, r)
        blanks = [COLUMNS[c] for c, v in r.items() if not v]
        if blanks:
            print(f"  Can't register: blank {', '.join(blanks)}. Skipping.")
            log_result(args.results, row_no, hex_id, "skipped", detail="blank: " + ", ".join(blanks))
            counts["skipped"] += 1
            continue
        if r["environment"].upper() not in SUPPORTED_ENVIRONMENTS:
            msg = (f"environment {r['environment']!r} not supported yet "
                   f"(supported: {', '.join(SUPPORTED_ENVIRONMENTS)})")
            print(f"  Can't register: {msg}. Skipping.")
            log_result(args.results, row_no, hex_id, "skipped", detail=msg)
            counts["skipped"] += 1
            continue

        while True:
            a = input("Register this beacon? [y]es / [s]kip / [q]uit: ").strip().lower()
            if a in ("y", "s", "q"):
                break
        if a == "q":
            print("Stopped.")
            break
        if a == "s":
            log_result(args.results, row_no, hex_id, "skipped", detail="skipped by user")
            counts["skipped"] += 1
            continue

        check_signed_in()
        try:
            status, ref, detail = fill_and_submit(s, to_config(r))
        except (StepFailed, ValueError) as e:
            status, ref, detail = "failed", "", str(e).replace("\n", " ")
        except requests.RequestException as e:
            status, ref, detail = "failed", "", f"site/network error before submit (nothing submitted): {e}"

        log_result(args.results, row_no, hex_id, status, ref, detail)
        if status == "registered":
            done[hex_id] = ref
            counts["registered"] += 1
            print(f"  Registered. Reference: {ref}" + (f"  ({detail})" if detail else ""))
        elif status == "UNKNOWN":
            print(f"  !! {detail}\n  Stopping so nothing gets registered twice.")
            break
        else:
            counts["failed"] += 1
            print(f"  Not registered: {detail}")

    print(f"\nDone. {counts['registered']} registered, {counts['skipped']} skipped, "
          f"{counts['failed']} failed. Details in {args.results}")


if __name__ == "__main__":
    main()
