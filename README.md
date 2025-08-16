# SageOGSync
Sync's the past 30 and future 30 days from your Outlook Client to your Google Calendar 
Step-by-Step Setup
1. Install Python and Required Packages

Make sure Python 3.10+ is installed.

Open a terminal/PowerShell and install required packages:

pip install pywin32 google-auth google-auth-oauthlib google-api-python-client requests


pywin32 → Access the local Outlook client

google-auth, google-auth-oauthlib, google-api-python-client → Access Google Calendar API

2. Prepare Google API Credentials

Go to Google Cloud Console.

Create a new project (or select an existing one).

Navigate to APIs & Services → Credentials → Create Credentials → OAuth client ID.

Choose Desktop app as the application type.

Download the JSON file and save it as:

google_client_secret.json


in the same folder as the Python script.

3. Enable Google Calendar API

In Google Cloud Console, go to APIs & Services → Library.

Search for Google Calendar API and enable it.

4. Prepare the Script

Save the Python script (the full GUI version) as OutlookToGoogleSync.py.

Make sure the following configuration variables are correct:

GOOGLE_CREDENTIALS_FILE = "google_client_secret.json"
SYNC_LOOKBACK_DAYS = 30
SYNC_LOOKAHEAD_DAYS = 30
SYNC_INTERVAL_MINUTES = 15


Ensure Python can run win32com.client (Outlook COM access).

Outlook must be installed and configured with your account on this Windows machine.

5. Run the Script

Open a terminal/PowerShell in the script’s folder.

Run:

python OutlookToGoogleSync.py


The script will:

Launch a GUI window for logging.

Prompt you to authenticate with Google the first time.

Start syncing events automatically every 15 minutes.

Step-by-Step Outline of What the Script Does
A. Google Calendar Setup

Loads credentials from google_token.json if it exists.

If not, it triggers OAuth to authenticate and save a new token.

Initializes a Google Calendar API service object.

B. Fetch Outlook Events

Initializes COM for thread safety.

Accesses Outlook Calendar folder.

Uses Restrict to fetch events within the past 30 days → next 30 days.

Includes recurring events.

Logs the number of events fetched and any skipped due to errors.

C. Fetch Google Events

Requests all Google Calendar events in the same 60-day window.

Maps Google events by OutlookEntryID (stored in extendedProperties.private).

D. Sync Logic

Loops through all Outlook events:

Checks if a Google event with the same OutlookEntryID exists.

If exists → updates the Google event.

If not → creates a new Google event.

Loops through Google events:

Deletes any Google event whose OutlookEntryID is no longer in Outlook within the 60-day window.

Logs each action:

Event created

Event updated

Event deleted

E. Background Sync

A separate thread runs the sync automatically every 15 minutes.

GUI logs display:

Sync start

Event counts

Detailed actions for each event

F. Duplicate Prevention

Duplicates are prevented by mapping Outlook EntryID → Google event.

Each Outlook event has a unique identifier stored in Google’s extendedProperties.

Only one Google event is created per Outlook event.

✅ Additional Notes

The GUI shows debug logs in real-time.

The sync is one-way: Outlook → Google.

Timezone-aware datetimes ensure accurate scheduling and comparisons.

Sync window is configurable with SYNC_LOOKBACK_DAYS and SYNC_LOOKAHEAD_DAYS.
