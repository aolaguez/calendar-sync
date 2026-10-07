# Google Calendar <-> Outlook Two-Way Sync

Run one script, it syncs both calendars, prints a summary of exactly what changed.
State lives in `sync_state.json` so re-runs are incremental (only real changes get pushed).

> **Setup note:** copy `sync_config.example.json` to `sync_config.json` and edit it for your own
> calendars. Credentials, tokens and `sync_state.json` are gitignored and never committed —
> you generate your own in Steps 1 and 2 below.


## One-time setup (~15 min total)

You need Python 3.9+ installed. Check with `python3 --version` (Mac/Linux) or `python --version` (Windows).

### Step 1: Google Cloud (Calendar API access)

1. Go to https://console.cloud.google.com/ and create a new project (or use an existing one).
2. Go to **APIs & Services -> Library**, search "Google Calendar API", click **Enable**.
3. Go to **APIs & Services -> OAuth consent screen**.
   - User type: **External** (unless you have a Workspace org, then Internal is fine).
   - Fill in app name (e.g. "My Calendar Sync"), your email for support/dev contact.
   - On the Scopes step, you can skip adding scopes here.
   - On **Test users**, add your own Google account email. This matters — without it you'll get an access-blocked error.
4. Go to **APIs & Services -> Credentials -> Create Credentials -> OAuth client ID**.
   - Application type: **Desktop app**.
   - Name it anything.
   - Click **Create**, then **Download JSON**.
5. Rename the downloaded file to `google_credentials.json` and place it in this same folder as `sync_calendars.py`.

### Step 2: Azure AD (Outlook / Microsoft Graph access)

1. Go to https://portal.azure.com/ -> **Microsoft Entra ID -> App registrations -> New registration**.
2. Name it anything (e.g. "Calendar Sync").
3. Supported account types: choose **Accounts in any organizational directory and personal Microsoft accounts** (unless your Outlook is a work/school account only — then pick the org-only option).
4. Redirect URI: leave blank (we use device code flow, no redirect needed).
5. Click **Register**. Copy the **Application (client) ID** shown on the overview page.
6. Go to **Authentication** (left sidebar) -> scroll to **Advanced settings** -> set **Allow public client flows** to **Yes** -> Save.
7. Go to **API permissions -> Add a permission -> Microsoft Graph -> Delegated permissions** -> search and add `Calendars.ReadWrite` -> **Add permissions**.
   (Admin consent isn't required for personal accounts; for work/school accounts your IT admin may need to approve it.)
8. Create a file named `outlook_credentials.json` in this folder with:

```json
{
  "client_id": "PASTE_YOUR_APPLICATION_CLIENT_ID_HERE",
  "tenant_id": "consumers"
}
```

Use `"tenant_id": "consumers"` for a personal Outlook/Hotmail account, `"organizations"` for a work/school account not tied to one tenant, or your specific tenant ID (found on the Azure app overview page) if you want to lock it to one org.

### Step 3: Run it

**Windows:** double-click `run_sync.bat`
**Mac/Linux:** open Terminal in this folder and run `chmod +x run_sync.sh` once, then double-click `run_sync.sh` (or run `./run_sync.sh`)

First run will:
- Install dependencies automatically into a local `venv` folder
- Open a browser for Google sign-in
- Print a Microsoft device-login code + URL — go to https://microsoft.com/devicelogin, enter the code, sign in

After that, tokens are cached (`token_google.json`, `token_outlook.json`) so future runs won't ask you to sign in again unless a token expires.

## Choosing which calendars sync

`sync_config.json` controls this. Each entry pairs one Outlook calendar with one Google calendar:

```json
{
  "pairs": [
    {"outlook": "Calendar",      "google": "Outlook Sync",  "create": ["google"]},
    {"outlook": "Google Sync",   "google": "primary",       "create": ["outlook"]},
    {"outlook": "Fall 26 UIUC",  "google": "UIUC Fall 26",  "create": ["google"]}
  ]
}
```

This is the **mirror model**: nothing is ever written into a main calendar. Each
main calendar is a source that mirrors into a dedicated calendar on the other
side, so anything that goes wrong can be hidden with a checkbox or deleted
wholesale without touching your real events.

- `"create"` lists which sides may be created if missing. A name that isn't
  listed there must already exist — a typo stops the run and prints the real
  calendar names instead of silently making an empty calendar.

- `"outlook"` must match an Outlook calendar name exactly. Get it wrong and the script stops and prints the list of real names rather than guessing.
- `"google": "primary"` means your main Google calendar. Any other name is looked up by title and **created if it doesn't exist**.
- Add a pair to sync more, delete a pair to stop syncing it. Removing a pair leaves already-synced events in place; it just stops tracking them.
- Outlook calendars marked read-only (subscribed feeds, holiday calendars) can be read but not written to. The script warns and syncs one-way.

Run `run_diagnose.bat` any time to see every calendar on both sides, how many events each holds, and what the current config would do. It's read-only.

## What it does each run

1. Pulls events from both calendars, roughly 7 days back to 180 days forward (adjustable — see `DAYS_PAST` / `DAYS_FUTURE` in `sync_calendars.py`).
2. For events it's already synced before: checks if either side changed, pushes the change to the other side.
3. New event on Google that isn't synced yet -> created on Outlook. Same in reverse.
4. Deleted on one side -> deleted on the other.
5. Edited on **both** sides since the last sync -> conflict. The more recently modified version wins; this is logged in the summary so you know it happened.
6. Prints a summary: counts of created/updated/deleted on each side, plus any errors.

## Known limitations (worth knowing before you rely on this)

- **Recurring events**: common patterns (daily/weekly/monthly/yearly, with interval/count/until) translate between Google's RRULE and Outlook's pattern format. Unusual recurrence rules (e.g. "the 2nd Tuesday of every month") aren't translated — the script skips creating the recurring copy on the other side rather than guessing wrong, and logs it as an error so you can see it happened.
- **Attendees/invites are not synced** — this syncs your own event data (title, time, location, description), not meeting invitations to other people. Syncing attendees would mean double-inviting people, which is its own can of worms.
- **All-day events** are supported. **Time zones** are normalized to UTC internally; double check display looks right in both apps after first sync.
- This is a **pull-based** script — it syncs when you run it, not continuously in real time. For near-real-time sync you'd want it running on a schedule (see below) or rebuilt as a hosted webhook-based service.

## Optional: run it automatically on a schedule

Instead of clicking it manually, you can schedule it:

- **Windows**: Task Scheduler -> create a task that runs `run_sync.bat` every N minutes.
- **Mac/Linux**: add a cron job, e.g. `*/30 * * * * /path/to/calendar_sync/run_sync.sh` for every 30 minutes.

## Files in this folder

| File | Purpose |
|---|---|
| `sync_calendars.py` | The actual sync logic |
| `sync_config.example.json` | Template — copy to `sync_config.json` and edit |
| `diagnose.py` / `run_diagnose.bat` | Read-only report on both sides — changes nothing |
| `requirements.txt` | Python dependencies |
| `run_sync.bat` / `run_sync.sh` | One-click launchers |
| `google_credentials.json` | You provide this (Step 1) |
| `outlook_credentials.json` | You provide this (Step 2) |
| `token_google.json` / `token_outlook.json` | Auto-created after first sign-in |
| `sync_state.json` | Auto-created — tracks which events are linked across calendars |
