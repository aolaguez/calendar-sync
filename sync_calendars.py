#!/usr/bin/env python3
"""
Two-way sync between Google Calendar and Outlook (Microsoft Graph) Calendar.

Run it, watch the summary print out, done. State is kept in sync_state.json
so re-runs only touch what actually changed.

See README.md for one-time setup (Google Cloud + Azure app registration).
"""

import os
import sys
import json
import hashlib
import datetime
import traceback

# ---- third-party libs (see requirements.txt) ----
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import msal
import requests

# ============================================================
# CONFIG
# ============================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

GOOGLE_CREDENTIALS_FILE = os.path.join(SCRIPT_DIR, "google_credentials.json")
GOOGLE_TOKEN_FILE = os.path.join(SCRIPT_DIR, "token_google.json")
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar"]

OUTLOOK_CONFIG_FILE = os.path.join(SCRIPT_DIR, "outlook_credentials.json")
OUTLOOK_TOKEN_CACHE_FILE = os.path.join(SCRIPT_DIR, "token_outlook.json")
OUTLOOK_SCOPES = ["Calendars.ReadWrite"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

STATE_FILE = os.path.join(SCRIPT_DIR, "sync_state.json")
CONFIG_FILE = os.path.join(SCRIPT_DIR, "sync_config.json")

# Which calendars to sync. Edit sync_config.json, not this.
# "google": "primary" means your main Google calendar. Any other name is
# looked up by title, and created if it doesn't exist yet.
#
# "create" lists which sides may be created if they don't exist. A name NOT
# listed there must already exist, and a typo stops the run instead of
# silently making an empty calendar and syncing nothing into it.
DEFAULT_CONFIG = {
    "pairs": [
        {"outlook": "Calendar", "google": "Outlook Sync", "create": ["google"]},
        {"outlook": "Google Sync", "google": "primary", "create": ["outlook"]},
        {"outlook": "Fall 26 UIUC", "google": "UIUC Fall 26", "create": ["google"]},
    ]
}

# If a calendar that previously held events suddenly looks empty, that's far
# more likely to be an API hiccup than a real mass deletion. Rather than
# propagate hundreds of deletions, bail out on that pair. Tuned as a fraction
# of the events we had mapped last run.
# What v1 of the state file implicitly described: the default Outlook calendar
# synced against Google primary. Used only when migrating an old state file.
LEGACY_V1_PAIR = {"outlook": "Calendar", "google": "primary"}

MASS_DELETE_ABORT_RATIO = 0.5
MASS_DELETE_ABORT_FLOOR = 10  # never abort on tiny calendars; deleting 3 of 4 events is normal

# How far back / forward to look each run. Wider window = safer but slower.
DAYS_PAST = 7
DAYS_FUTURE = 400

# A marker we stamp on every event we create, so we never re-import
# our own copies as "new" events.
SYNC_MARKER_KEY = "calendar_sync_tool_id"


# ============================================================
# AUTH
# ============================================================

def auth_google():
    """Returns an authenticated Google Calendar API service object."""
    creds = None
    if os.path.exists(GOOGLE_TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE, GOOGLE_SCOPES)

    if not creds or not creds.valid:
        refreshed = False
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                refreshed = True
            except RefreshError as e:
                # Refresh token is dead (revoked, password change, or the 7-day
                # expiry that applies while the OAuth app is in "Testing" status).
                # Fall back to a fresh browser sign-in instead of crashing.
                print(f"Saved Google login is no longer valid ({e}).")
                print("Opening a browser to sign in again...")
                creds = None

        if not refreshed:
            if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
                print(f"ERROR: Missing {GOOGLE_CREDENTIALS_FILE}")
                print("Follow the Google setup steps in README.md first.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(
                GOOGLE_CREDENTIALS_FILE, GOOGLE_SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(GOOGLE_TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)


def auth_outlook():
    """Returns a valid Microsoft Graph access token string."""
    if not os.path.exists(OUTLOOK_CONFIG_FILE):
        print(f"ERROR: Missing {OUTLOOK_CONFIG_FILE}")
        print("Follow the Azure setup steps in README.md first.")
        sys.exit(1)

    with open(OUTLOOK_CONFIG_FILE) as f:
        cfg = json.load(f)

    cache = msal.SerializableTokenCache()
    if os.path.exists(OUTLOOK_TOKEN_CACHE_FILE):
        cache.deserialize(open(OUTLOOK_TOKEN_CACHE_FILE, "r").read())

    app = msal.PublicClientApplication(
        cfg["client_id"],
        authority=f"https://login.microsoftonline.com/{cfg.get('tenant_id', 'consumers')}",
        token_cache=cache,
    )

    accounts = app.get_accounts()
    result = None
    if accounts:
        try:
            result = app.acquire_token_silent(OUTLOOK_SCOPES, account=accounts[0])
        except Exception as e:
            # Corrupt cache, clock skew, transient network — fall through to
            # an interactive sign-in rather than dying here.
            print(f"Could not reuse saved Microsoft login ({e}).")
            result = None

    # acquire_token_silent returns None when it needs interaction, but can also
    # return a dict carrying an error (e.g. the refresh token was revoked or
    # aged out past 90 days). Treat anything without a token as "sign in again".
    if result and "access_token" not in result:
        print("Saved Microsoft login is no longer valid "
              f"({result.get('error_description', result.get('error', 'unknown error'))}).")
        result = None

    if not result:
        flow = app.initiate_device_flow(scopes=OUTLOOK_SCOPES)
        if "user_code" not in flow:
            print("ERROR: Could not start Microsoft device login flow.")
            print(flow)
            sys.exit(1)
        print("\n--- Microsoft sign-in required ---")
        print(flow["message"])
        print("-----------------------------------\n")
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        with open(OUTLOOK_TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())

    if "access_token" not in result:
        print("ERROR: Microsoft auth failed.")
        print(result.get("error_description", result))
        sys.exit(1)

    return result["access_token"]


# ============================================================
# STATE (the mapping between the two calendars)
# ============================================================

def load_config():
    """Which calendar pairs to sync. Writes a starter file on first run."""
    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w") as f:
            json.dump(DEFAULT_CONFIG, f, indent=2)
        print(f"Created {os.path.basename(CONFIG_FILE)} with default calendar pairs.")
        return DEFAULT_CONFIG

    with open(CONFIG_FILE) as f:
        cfg = json.load(f)

    pairs = cfg.get("pairs")
    if not pairs:
        print(f"ERROR: {os.path.basename(CONFIG_FILE)} has no 'pairs' entry.")
        sys.exit(1)
    for p in pairs:
        if not p.get("outlook") or not p.get("google"):
            print(f"ERROR: every pair needs both 'outlook' and 'google'. Bad entry: {p}")
            sys.exit(1)
    return cfg


def pair_key(pair):
    return f"{pair['outlook']}=>{pair['google']}"


def make_sync_id(pair, google_id, outlook_id):
    """
    Stable, collision-free key for one linked pair of events.

    The old scheme was f"sync_{id[:12]}", but every Microsoft Graph event ID in
    a mailbox shares a long common prefix, so every Outlook-origin event
    collapsed onto the same key and overwrote the previous one. Hash the full
    IDs instead.
    """
    blob = f"{pair_key(pair)}|{google_id or ''}|{outlook_id or ''}"
    return hashlib.sha1(blob.encode("utf-8")).hexdigest()[:20]


def load_state():
    """
    State layout (v2):
      {"version": 2, "pairs": {"<pair key>": {"mappings": {...}}}, "last_sync": ...}

    v1 was a single flat "mappings" dict covering the default calendar only.
    It gets migrated into the first pair so nothing re-syncs or duplicates.
    """
    if not os.path.exists(STATE_FILE):
        return {"version": 2, "pairs": {}}

    with open(STATE_FILE) as f:
        state = json.load(f)

    if state.get("version") == 2:
        state.setdefault("pairs", {})
        return state

    # --- migrate v1 -> v2 ---
    # v1 only ever synced the default Outlook calendar against Google primary,
    # so that is the pair those links belong to. This must NOT follow whatever
    # happens to be first in the current config — the links describe events
    # that already exist in those two specific calendars.
    legacy = state.get("mappings", {})
    migrated_key = pair_key(LEGACY_V1_PAIR)
    new_state = {
        "version": 2,
        "pairs": {migrated_key: {"mappings": legacy}},
        "last_sync": state.get("last_sync"),
    }
    if legacy:
        print(f"Migrated {len(legacy)} existing event link(s) into pair '{migrated_key}'.")
    return new_state


def save_state(state):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


def pair_mappings(state, pair):
    """The mappings dict for one pair, created on demand."""
    return state["pairs"].setdefault(pair_key(pair), {"mappings": {}})["mappings"]


# ============================================================
# EVENT NORMALIZATION
# A common shape both APIs get translated into/out of, so the
# diff/sync logic only has to deal with one format.
# ============================================================

def compute_hash(norm):
    """Hash of the fields that matter for detecting a real change."""
    key_fields = {
        "title": norm.get("title", ""),
        "description": norm.get("description", ""),
        "location": norm.get("location", ""),
        "start": norm.get("start", ""),
        "end": norm.get("end", ""),
        "all_day": norm.get("all_day", False),
        "recurrence": norm.get("recurrence", ""),
    }
    blob = json.dumps(key_fields, sort_keys=True)
    return hashlib.md5(blob.encode("utf-8")).hexdigest()


def google_to_norm(ev):
    start = ev["start"].get("dateTime") or ev["start"].get("date")
    end = ev["end"].get("dateTime") or ev["end"].get("date")
    all_day = "date" in ev["start"]
    recurrence = ";".join(ev.get("recurrence", [])) if ev.get("recurrence") else ""
    return {
        "title": ev.get("summary", "(no title)"),
        "description": ev.get("description", ""),
        "location": ev.get("location", ""),
        "start": start,
        "end": end,
        "all_day": all_day,
        "recurrence": recurrence,
        "last_modified": ev.get("updated"),
        "source_id": ev["id"],
    }


def norm_to_google(norm):
    body = {
        "summary": norm["title"],
        "description": norm.get("description", ""),
        "location": norm.get("location", ""),
    }
    if norm["all_day"]:
        body["start"] = {"date": norm["start"][:10]}
        body["end"] = {"date": norm["end"][:10]}
    else:
        body["start"] = {"dateTime": norm["start"]}
        body["end"] = {"dateTime": norm["end"]}
    if norm.get("recurrence"):
        body["recurrence"] = norm["recurrence"].split(";")
    return body


def outlook_to_norm(ev):
    start = ev["start"]["dateTime"] + "Z" if "Z" not in ev["start"]["dateTime"] else ev["start"]["dateTime"]
    end = ev["end"]["dateTime"] + "Z" if "Z" not in ev["end"]["dateTime"] else ev["end"]["dateTime"]
    all_day = ev.get("isAllDay", False)
    recurrence = outlook_pattern_to_rrule(ev.get("recurrence")) if ev.get("recurrence") else ""
    return {
        "title": ev.get("subject", "(no title)"),
        "description": (ev.get("body") or {}).get("content", ""),
        "location": (ev.get("location") or {}).get("displayName", ""),
        "start": start,
        "end": end,
        "all_day": all_day,
        "recurrence": recurrence,
        "last_modified": ev.get("lastModifiedDateTime"),
        "source_id": ev["id"],
    }


def norm_to_outlook(norm):
    body = {
        "subject": norm["title"],
        "body": {"contentType": "text", "content": norm.get("description", "")},
        "location": {"displayName": norm.get("location", "")},
        "isAllDay": norm["all_day"],
        "start": {"dateTime": norm["start"], "timeZone": "UTC"},
        "end": {"dateTime": norm["end"], "timeZone": "UTC"},
    }
    if norm.get("recurrence"):
        pattern = rrule_to_outlook_pattern(norm["recurrence"], norm["start"])
        if pattern:
            body["recurrence"] = pattern
    return body


# ============================================================
# RECURRENCE TRANSLATION (best-effort: DAILY/WEEKLY/MONTHLY/YEARLY,
# with INTERVAL/COUNT/UNTIL/BYDAY). Complex RRULEs pass through as
# a plain (non-recurring) event rather than crashing the sync.
# ============================================================

WEEKDAY_MAP = {
    "MO": "monday", "TU": "tuesday", "WE": "wednesday", "TH": "thursday",
    "FR": "friday", "SA": "saturday", "SU": "sunday",
}
WEEKDAY_MAP_REV = {v: k for k, v in WEEKDAY_MAP.items()}


def rrule_to_outlook_pattern(rrule_str, start_iso):
    try:
        parts = {}
        for chunk in rrule_str.replace("RRULE:", "").split(";"):
            if "=" in chunk:
                k, v = chunk.split("=", 1)
                parts[k] = v

        freq = parts.get("FREQ", "")
        interval = int(parts.get("INTERVAL", 1))
        freq_map = {"DAILY": "daily", "WEEKLY": "weekly", "MONTHLY": "absoluteMonthly", "YEARLY": "absoluteYearly"}
        if freq not in freq_map:
            return None  # unsupported pattern, fall back to one-off event

        pattern = {"type": freq_map[freq], "interval": interval}
        if freq == "WEEKLY" and "BYDAY" in parts:
            pattern["daysOfWeek"] = [WEEKDAY_MAP[d] for d in parts["BYDAY"].split(",") if d in WEEKDAY_MAP]
        if freq == "MONTHLY":
            pattern["dayOfMonth"] = int(start_iso[8:10])
        if freq == "YEARLY":
            pattern["dayOfMonth"] = int(start_iso[8:10])
            pattern["month"] = int(start_iso[5:7])

        rng = {"type": "noEnd", "startDate": start_iso[:10]}
        if "COUNT" in parts:
            rng = {"type": "numbered", "startDate": start_iso[:10], "numberOfOccurrences": int(parts["COUNT"])}
        elif "UNTIL" in parts:
            until = parts["UNTIL"][:8]
            rng = {"type": "endDate", "startDate": start_iso[:10],
                   "endDate": f"{until[0:4]}-{until[4:6]}-{until[6:8]}"}

        return {"pattern": pattern, "range": rng}
    except Exception:
        return None


def outlook_pattern_to_rrule(rec):
    try:
        pattern = rec.get("pattern", {})
        rng = rec.get("range", {})
        type_map = {"daily": "DAILY", "weekly": "WEEKLY", "absoluteMonthly": "MONTHLY", "absoluteYearly": "YEARLY"}
        freq = type_map.get(pattern.get("type"))
        if not freq:
            return ""

        rule = [f"FREQ={freq}", f"INTERVAL={pattern.get('interval', 1)}"]
        if pattern.get("daysOfWeek"):
            days = [WEEKDAY_MAP_REV[d] for d in pattern["daysOfWeek"] if d in WEEKDAY_MAP_REV]
            if days:
                rule.append(f"BYDAY={','.join(days)}")
        if rng.get("type") == "numbered":
            rule.append(f"COUNT={rng.get('numberOfOccurrences', 1)}")
        elif rng.get("type") == "endDate":
            end_date = rng.get("endDate", "").replace("-", "")
            rule.append(f"UNTIL={end_date}T000000Z")

        return "RRULE:" + ";".join(rule)
    except Exception:
        return ""


# ============================================================
# GOOGLE CALENDAR API HELPERS
# ============================================================

def google_resolve_calendar(service, name, may_create):
    """
    Turn a calendar name from the config into a Google calendar ID.
    "primary" is passed straight through. Anything else is matched by title.
    Created only if this pair's "create" list says so.
    """
    if name == "primary":
        return "primary"

    names = []
    page_token = None
    while True:
        resp = service.calendarList().list(pageToken=page_token).execute(num_retries=5)
        for c in resp.get("items", []):
            names.append(c.get("summary"))
            if c.get("summary") == name:
                return c["id"]
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    if not may_create:
        print(f"ERROR: no Google calendar named '{name}', and this pair isn't "
              f"allowed to create one.")
        print("Google calendars on this account:")
        for n in names:
            print(f"  - {n}")
        print(f"\nFix the name in {os.path.basename(CONFIG_FILE)}, or add "
              f"\"create\": [\"google\"] to that pair.")
        sys.exit(1)

    created = service.calendars().insert(body={"summary": name}).execute(num_retries=5)
    print(f"  Created new Google calendar: {name}")
    return created["id"]


def google_list_events(service, calendar_id, time_min, time_max):
    events = []
    page_token = None
    while True:
        resp = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min, timeMax=time_max,
            singleEvents=False,  # keep recurring masters as one entry
            pageToken=page_token,
            maxResults=2500,
        ).execute(num_retries=5)
        events.extend(resp.get("items", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    # Drop cancelled instances
    return [e for e in events if e.get("status") != "cancelled"]


def google_create_event(service, calendar_id, norm):
    body = norm_to_google(norm)
    created = service.events().insert(calendarId=calendar_id, body=body).execute(num_retries=5)
    return created["id"]


def google_update_event(service, calendar_id, event_id, norm):
    body = norm_to_google(norm)
    service.events().patch(
        calendarId=calendar_id, eventId=event_id, body=body
    ).execute(num_retries=5)


def google_delete_event(service, calendar_id, event_id):
    try:
        service.events().delete(
            calendarId=calendar_id, eventId=event_id
        ).execute(num_retries=5)
    except Exception as e:
        if "404" not in str(e) and "410" not in str(e):
            raise


# ============================================================
# OUTLOOK (MICROSOFT GRAPH) API HELPERS
# ============================================================

def _graph_headers(token):
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def outlook_resolve_calendar(token, name, may_create):
    """
    Turn a calendar name from the config into an Outlook calendar ID.
    Created only if this pair's "create" list says so — otherwise a typo
    stops the run rather than silently making an empty calendar.
    """
    resp = requests.get(f"{GRAPH_BASE}/me/calendars?$top=100", headers=_graph_headers(token))
    resp.raise_for_status()
    cals = resp.json().get("value", [])

    for c in cals:
        if c.get("name") == name:
            if not c.get("canEdit", True):
                print(f"  NOTE: Outlook calendar '{name}' is read-only "
                      f"(subscribed or shared). Changes can't be pushed back to it.")
            return c["id"]

    if not may_create:
        print(f"ERROR: no Outlook calendar named '{name}', and this pair isn't "
              f"allowed to create one.")
        print("Calendars on this account:")
        for c in cals:
            print(f"  - {c.get('name')}")
        print(f"\nFix the name in {os.path.basename(CONFIG_FILE)}, or add "
              f"\"create\": [\"outlook\"] to that pair.")
        sys.exit(1)

    resp = requests.post(f"{GRAPH_BASE}/me/calendars",
                         headers=_graph_headers(token), json={"name": name})
    resp.raise_for_status()
    print(f"  Created new Outlook calendar: {name}")
    return resp.json()["id"]


def outlook_list_events(token, calendar_id, time_min, time_max):
    events = []
    url = (f"{GRAPH_BASE}/me/calendars/{calendar_id}/calendarView"
           f"?startDateTime={time_min}&endDateTime={time_max}"
           f"&$top=250")
    headers = _graph_headers(token)
    headers["Prefer"] = 'outlook.timezone="UTC"'
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data = resp.json()
        events.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return events


def outlook_create_event(token, calendar_id, norm):
    body = norm_to_outlook(norm)
    resp = requests.post(f"{GRAPH_BASE}/me/calendars/{calendar_id}/events",
                         headers=_graph_headers(token), json=body)
    resp.raise_for_status()
    return resp.json()["id"]


def outlook_update_event(token, event_id, norm):
    body = norm_to_outlook(norm)
    resp = requests.patch(f"{GRAPH_BASE}/me/events/{event_id}", headers=_graph_headers(token), json=body)
    resp.raise_for_status()


def outlook_delete_event(token, event_id):
    resp = requests.delete(f"{GRAPH_BASE}/me/events/{event_id}", headers=_graph_headers(token))
    if resp.status_code not in (204, 404):
        resp.raise_for_status()


# ============================================================
# SYNC ORCHESTRATION
# ============================================================

def check_mass_delete(side, name, seen_count, mapped_count):
    """
    Guard against propagating a mass deletion that probably didn't happen.

    The reconciliation logic reads "mapped event is missing from the fetch" as
    "the user deleted it, delete the other copy too". If a fetch fails partway,
    or a calendar is momentarily empty, that logic would happily wipe hundreds
    of real events off the other side. Refuse to run the pair instead.
    """
    if mapped_count < MASS_DELETE_ABORT_FLOOR:
        return None
    missing = mapped_count - seen_count
    if missing > mapped_count * MASS_DELETE_ABORT_RATIO:
        return (f"{side} calendar '{name}' returned {seen_count} events but "
                f"{mapped_count} were synced last run. That looks like a failed "
                f"fetch rather than {missing} real deletions, so this pair was "
                f"skipped instead of deleting them from the other calendar. "
                f"Re-run; if it persists and the deletions are real, delete the "
                f"pair's section from sync_state.json to reset it.")
    return None


def sync_pair(g_service, o_token, pair, state, time_min, time_max, stats):
    """Sync one Outlook calendar against one Google calendar."""
    label = f"{pair['outlook']}  <->  {pair['google']}"
    print(f"\n--- {label} ---")

    creatable = pair.get("create", [])
    g_cal = google_resolve_calendar(g_service, pair["google"], "google" in creatable)
    o_cal = outlook_resolve_calendar(o_token, pair["outlook"], "outlook" in creatable)

    g_events_raw = google_list_events(g_service, g_cal, time_min, time_max)
    o_events_raw = outlook_list_events(o_token, o_cal,
                                       time_min.replace(".000Z", ""),
                                       time_max.replace(".000Z", ""))

    g_events = {e["id"]: google_to_norm(e) for e in g_events_raw}
    o_events = {e["id"]: outlook_to_norm(e) for e in o_events_raw}
    print(f"  Google: {len(g_events)} events   Outlook: {len(o_events)} events")

    mappings = pair_mappings(state, pair)

    mapped_g_ids = {m["google_id"] for m in mappings.values() if m.get("google_id")}
    mapped_o_ids = {m["outlook_id"] for m in mappings.values() if m.get("outlook_id")}

    # Safety check before any destructive work.
    for side, name, events, mapped in (
        ("Google", pair["google"], g_events, mapped_g_ids),
        ("Outlook", pair["outlook"], o_events, mapped_o_ids),
    ):
        still_present = len([i for i in mapped if i in events])
        problem = check_mass_delete(side, name, still_present, len(mapped))
        if problem:
            print(f"  SKIPPED: {problem}")
            stats["errors"].append(problem)
            return

    # ---- 1. Walk existing mappings: update / delete propagation ----
    for sync_id in list(mappings.keys()):
        m = mappings[sync_id]
        g_id, o_id = m.get("google_id"), m.get("outlook_id")
        g_ev = g_events.get(g_id) if g_id else None
        o_ev = o_events.get(o_id) if o_id else None

        try:
            if g_ev is None and o_ev is None:
                del mappings[sync_id]
                continue

            if g_ev is None and o_ev is not None:
                outlook_delete_event(o_token, o_id)
                stats["deleted_outlook"] += 1
                del mappings[sync_id]
                continue

            if o_ev is None and g_ev is not None:
                google_delete_event(g_service, g_cal, g_id)
                stats["deleted_google"] += 1
                del mappings[sync_id]
                continue

            g_hash = compute_hash(g_ev)
            o_hash = compute_hash(o_ev)
            g_changed = g_hash != m.get("last_hash_g")
            o_changed = o_hash != m.get("last_hash_o")

            if g_changed and o_changed:
                # conflict: most recently modified side wins
                stats["conflicts"] += 1
                g_time = g_ev.get("last_modified") or ""
                o_time = o_ev.get("last_modified") or ""
                if g_time >= o_time:
                    outlook_update_event(o_token, o_id, g_ev)
                    stats["updated_outlook"] += 1
                else:
                    google_update_event(g_service, g_cal, g_id, o_ev)
                    stats["updated_google"] += 1
            elif g_changed:
                outlook_update_event(o_token, o_id, g_ev)
                stats["updated_outlook"] += 1
            elif o_changed:
                google_update_event(g_service, g_cal, g_id, o_ev)
                stats["updated_google"] += 1

            mappings[sync_id]["last_hash_g"] = compute_hash(g_ev)
            mappings[sync_id]["last_hash_o"] = compute_hash(o_ev)

        except Exception as e:
            stats["errors"].append(f"sync_id {sync_id}: {e}")

    # ---- 2. New Google events not yet mapped -> create in Outlook ----
    to_create_outlook = [(i, e) for i, e in g_events.items() if i not in mapped_g_ids]
    for n, (g_id, g_ev) in enumerate(to_create_outlook, 1):
        try:
            o_id = outlook_create_event(o_token, o_cal, g_ev)
            sync_id = make_sync_id(pair, g_id, o_id)
            mappings[sync_id] = {
                "google_id": g_id, "outlook_id": o_id,
                "last_hash_g": compute_hash(g_ev), "last_hash_o": compute_hash(g_ev),
            }
            stats["created_outlook"] += 1
            if n % 25 == 0:
                print(f"    ...created {n}/{len(to_create_outlook)} on Outlook")
        except Exception as e:
            stats["errors"].append(f"creating outlook copy of google event {g_id}: {e}")

    # ---- 3. New Outlook events not yet mapped -> create in Google ----
    to_create_google = [(i, e) for i, e in o_events.items() if i not in mapped_o_ids]
    for n, (o_id, o_ev) in enumerate(to_create_google, 1):
        try:
            g_id = google_create_event(g_service, g_cal, o_ev)
            sync_id = make_sync_id(pair, g_id, o_id)
            mappings[sync_id] = {
                "google_id": g_id, "outlook_id": o_id,
                "last_hash_g": compute_hash(o_ev), "last_hash_o": compute_hash(o_ev),
            }
            stats["created_google"] += 1
            if n % 25 == 0:
                print(f"    ...created {n}/{len(to_create_google)} on Google")
        except Exception as e:
            stats["errors"].append(f"creating google copy of outlook event {o_id}: {e}")


def run_sync():
    now = datetime.datetime.now(datetime.timezone.utc)
    time_min = (now - datetime.timedelta(days=DAYS_PAST)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    time_max = (now + datetime.timedelta(days=DAYS_FUTURE)).strftime("%Y-%m-%dT%H:%M:%S.000Z")

    config = load_config()

    print("Signing in to Google Calendar...")
    g_service = auth_google()
    print("Signing in to Outlook...")
    o_token = auth_outlook()

    state = load_state()

    stats = {"created_outlook": 0, "created_google": 0,
             "updated_outlook": 0, "updated_google": 0,
             "deleted_outlook": 0, "deleted_google": 0,
             "conflicts": 0, "errors": []}

    for pair in config["pairs"]:
        try:
            sync_pair(g_service, o_token, pair, state, time_min, time_max, stats)
        except SystemExit:
            raise
        except Exception as e:
            stats["errors"].append(f"pair {pair_key(pair)}: {e}")
            print(f"  FAILED: {e}")
        finally:
            # Save after every pair so a later failure can't lose earlier work.
            state["last_sync"] = now.isoformat()
            save_state(state)

    return stats


def print_summary(stats):
    print("\n" + "=" * 50)
    print("SYNC COMPLETE")
    print("=" * 50)
    print(f"  Created on Outlook : {stats['created_outlook']}")
    print(f"  Created on Google  : {stats['created_google']}")
    print(f"  Updated on Outlook : {stats['updated_outlook']}")
    print(f"  Updated on Google  : {stats['updated_google']}")
    print(f"  Deleted on Outlook : {stats['deleted_outlook']}")
    print(f"  Deleted on Google  : {stats['deleted_google']}")
    print(f"  Conflicts resolved : {stats['conflicts']} (most recently edited side won)")
    if stats["errors"]:
        print(f"\n  {len(stats['errors'])} ERROR(S):")
        for err in stats["errors"][:20]:
            print(f"    - {err}")
        if len(stats["errors"]) > 20:
            print(f"    ... and {len(stats['errors']) - 20} more")
        print("\nSTATUS: COMPLETED WITH ERRORS")
    else:
        print("\nSTATUS: SUCCESS - both calendars are in sync")
    print("=" * 50)


if __name__ == "__main__":
    try:
        stats = run_sync()
        print_summary(stats)
        sys.exit(0 if not stats["errors"] else 1)
    except Exception:
        print("\nFATAL ERROR - sync did not complete:")
        traceback.print_exc()
        sys.exit(1)
