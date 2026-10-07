"""
Read-only diagnostic for the calendar sync.

Creates nothing, changes nothing, deletes nothing. It just reports what the
sync script can actually SEE on each side, so we can find out why the semester
schedule isn't crossing over.

Run it the same way you run the sync: double-click run_diagnose.bat
"""

import datetime
import collections
import requests

import sync_calendars as sc


def main():
    now = datetime.datetime.now(datetime.timezone.utc)
    time_min = (now - datetime.timedelta(days=sc.DAYS_PAST)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    time_max = (now + datetime.timedelta(days=sc.DAYS_FUTURE)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    o_min = time_min.replace(".000Z", "")
    o_max = time_max.replace(".000Z", "")

    print(f"Window: {time_min}  ->  {time_max}\n")

    # ---------------- GOOGLE ----------------
    print("=" * 60)
    print("GOOGLE")
    print("=" * 60)
    g = sc.auth_google()

    cal_list = g.calendarList().list().execute().get("items", [])
    print(f"Calendars visible to this Google login: {len(cal_list)}")
    for c in cal_list:
        primary = "  <-- PRIMARY (the one the sync writes to)" if c.get("primary") else ""
        print(f"  - {c.get('summary')}  [{c.get('id')}]  access={c.get('accessRole')}{primary}")

    g_events = sc.google_list_events(g, "primary", time_min, time_max)
    print(f"\nEvents on primary calendar in window: {len(g_events)}")

    titles = collections.Counter(
        (e.get("summary", "(no title)"), (e.get("start") or {}).get("dateTime")
         or (e.get("start") or {}).get("date"))
        for e in g_events
    )
    dupes = {k: v for k, v in titles.items() if v > 1}
    if dupes:
        print(f"\n  !! {len(dupes)} title+time combos appear MORE THAN ONCE (duplicates):")
        for (t, s), n in list(dupes.items())[:15]:
            print(f"     {n}x  {t}  @ {s}")
        if len(dupes) > 15:
            print(f"     ... and {len(dupes) - 15} more")

    print("\n  Sample of what's there now:")
    for e in g_events[:10]:
        s = (e.get("start") or {}).get("dateTime") or (e.get("start") or {}).get("date")
        print(f"     {s}  {e.get('summary')}")

    # ---------------- OUTLOOK ----------------
    print("\n" + "=" * 60)
    print("OUTLOOK")
    print("=" * 60)
    token = sc.auth_outlook()
    headers = sc._graph_headers(token)
    headers["Prefer"] = 'outlook.timezone="UTC"'

    me = requests.get(f"{sc.GRAPH_BASE}/me", headers=sc._graph_headers(token)).json()
    print(f"Signed in as: {me.get('userPrincipalName')}\n")

    cals = requests.get(f"{sc.GRAPH_BASE}/me/calendars?$top=100",
                        headers=sc._graph_headers(token)).json().get("value", [])
    print(f"Outlook calendars on this account: {len(cals)}")

    for c in cals:
        cid = c["id"]
        default = "  <-- DEFAULT (the ONLY one the sync reads)" if c.get("isDefaultCalendar") else ""
        try:
            r = requests.get(
                f"{sc.GRAPH_BASE}/me/calendars/{cid}/calendarView"
                f"?startDateTime={o_min}&endDateTime={o_max}&$top=250",
                headers=headers,
            )
            r.raise_for_status()
            evs = r.json().get("value", [])
            count = len(evs)
            more = "+" if r.json().get("@odata.nextLink") else ""
        except Exception as e:
            count, more, evs = f"ERROR: {e}", "", []

        print(f"\n  - {c.get('name')}{default}")
        print(f"      canEdit={c.get('canEdit')}  events in window: {count}{more}")
        for e in evs[:6]:
            s = (e.get("start") or {}).get("dateTime")
            print(f"        {s}  {e.get('subject')}  (type={e.get('type')})")
        if len(evs) > 6:
            print(f"        ... and {len(evs) - 6} more")

    # ---------------- CONFIG + STATE ----------------
    print("\n" + "=" * 60)
    print("CONFIGURED PAIRS")
    print("=" * 60)
    config = sc.load_config()
    state = sc.load_state()

    for pair in config["pairs"]:
        key = sc.pair_key(pair)
        mappings = state.get("pairs", {}).get(key, {}).get("mappings", {})
        print(f"\n  {key}")
        print(f"      linked events tracked in state: {len(mappings)}")

        try:
            o_cal = sc.outlook_resolve_calendar(token, pair["outlook"])
            o_events = sc.outlook_list_events(token, o_cal, o_min, o_max)
            o_ids = {m["outlook_id"] for m in mappings.values() if m.get("outlook_id")}
            print(f"      events on the Outlook side: {len(o_events)}")
            print(f"      not yet linked (would be created on Google): "
                  f"{len([e for e in o_events if e['id'] not in o_ids])}")
        except SystemExit:
            print("      ^ that calendar name doesn't exist — fix sync_config.json")


if __name__ == "__main__":
    main()
