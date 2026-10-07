"""
One-time migration: main-to-main sync  ->  mirror model.

Why this exists
---------------
The old setup synced Outlook "Calendar" directly against Google "primary", so
the same 10 events now sit in BOTH main calendars. Under the mirror model each
main calendar also gets mirrored to a dedicated calendar on the other side. If
we just switched over, every one of those events would end up in four places:
its two current homes plus two fresh mirrors.

This script removes the duplicate copies the old sync created, leaving each
event only in the calendar it originally came from. After that, a normal sync
run builds clean mirrors.

It shows you exactly what it will delete and asks before touching anything.
Run it once, then run the normal sync.
"""

import sys

import sync_calendars as sc

LEGACY_PAIR = sc.LEGACY_V1_PAIR


def describe(norm):
    return f"{norm.get('start', '?')[:16]}  {norm.get('title', '(no title)')}"


def main():
    state = sc.load_state()
    key = sc.pair_key(LEGACY_PAIR)
    mappings = state.get("pairs", {}).get(key, {}).get("mappings", {})

    if not mappings:
        print(f"No legacy '{key}' links found — nothing to migrate.")
        print("You can go straight to run_sync.bat.")
        return

    print(f"Found {len(mappings)} event(s) linked by the old main-to-main sync.\n")

    print("Signing in...")
    g_service = sc.auth_google()
    o_token = sc.auth_outlook()

    now = sc.datetime.datetime.now(sc.datetime.timezone.utc)
    time_min = (now - sc.datetime.timedelta(days=sc.DAYS_PAST)).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    time_max = (now + sc.datetime.timedelta(days=sc.DAYS_FUTURE)).strftime("%Y-%m-%dT%H:%M:%S.000Z")

    o_cal = sc.outlook_resolve_calendar(o_token, "Calendar", False)

    g_events = {e["id"]: sc.google_to_norm(e)
                for e in sc.google_list_events(g_service, "primary", time_min, time_max)}
    o_events = {e["id"]: sc.outlook_to_norm(e)
                for e in sc.outlook_list_events(o_token, o_cal,
                                                time_min.replace(".000Z", ""),
                                                time_max.replace(".000Z", ""))}

    # The old sync_id encoded which side an event came from: Google-origin
    # entries were keyed on the Google ID, Outlook-origin on the Outlook ID.
    # The copy to remove is the one on the OTHER side from the original.
    delete_from_google = []
    delete_from_outlook = []
    skipped = []

    for sync_id, m in mappings.items():
        g_id, o_id = m.get("google_id"), m.get("outlook_id")
        google_origin = bool(g_id) and sync_id == f"sync_{g_id[:12]}"

        # The original is the side the event came from; the copy is the other.
        # NEVER delete the copy unless the original is confirmed still present —
        # otherwise an event whose original was deleted by hand would lose its
        # last remaining copy and vanish entirely.
        if google_origin:
            original_alive = g_id in g_events
            copy_alive = o_id in o_events
            if copy_alive and original_alive:
                delete_from_outlook.append((o_id, o_events[o_id]))
            elif copy_alive and not original_alive:
                skipped.append(("Outlook", o_events[o_id]))
        else:
            original_alive = o_id in o_events
            copy_alive = g_id in g_events
            if copy_alive and original_alive:
                delete_from_google.append((g_id, g_events[g_id]))
            elif copy_alive and not original_alive:
                skipped.append(("Google", g_events[g_id]))

    if skipped:
        print(f"KEEPING {len(skipped)} event(s) — the original was already gone, so "
              f"this copy is the only one left:\n")
        for side, n in skipped:
            print(f"    - [{side}] {describe(n)}")
        print()

    if not delete_from_google and not delete_from_outlook:
        print("The linked events are already gone from both sides. Nothing to delete.")
    else:
        print("These are COPIES the old sync created. The original of each stays put.\n")
        if delete_from_outlook:
            print(f"  Delete from Outlook 'Calendar'  ({len(delete_from_outlook)}):")
            for _, n in delete_from_outlook:
                print(f"    - {describe(n)}")
        if delete_from_google:
            print(f"\n  Delete from Google 'primary'  ({len(delete_from_google)}):")
            for _, n in delete_from_google:
                print(f"    - {describe(n)}")

        print(f"\nTotal to delete: {len(delete_from_google) + len(delete_from_outlook)}")
        answer = input("\nProceed? Type yes to continue, anything else to abort: ").strip().lower()
        if answer != "yes":
            print("Aborted. Nothing was changed.")
            return

        for o_id, n in delete_from_outlook:
            try:
                sc.outlook_delete_event(o_token, o_id)
                print(f"  removed from Outlook: {describe(n)}")
            except Exception as e:
                print(f"  FAILED on {describe(n)}: {e}")

        for g_id, n in delete_from_google:
            try:
                sc.google_delete_event(g_service, "primary", g_id)
                print(f"  removed from Google: {describe(n)}")
            except Exception as e:
                print(f"  FAILED on {describe(n)}: {e}")

    del state["pairs"][key]
    sc.save_state(state)
    print(f"\nDropped the legacy '{key}' links from the state file.")
    print("Now run run_sync.bat — it will build the mirror calendars from scratch.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nAborted.")
        sys.exit(1)
