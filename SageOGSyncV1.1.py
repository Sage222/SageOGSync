import datetime
import threading
import time
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import os
import pythoncom
import queue
import pytz

import win32com.client
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# =========================================
# INSTRUCTIONS
# =========================================
# 1. Install required libraries:
#    pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib pywin32 pytz
#
# 2. Enable the Google Calendar API in your Google Cloud project:
#    https://console.cloud.google.com/apis/library/calendar-json.googleapis.com
#
# 3. Download your OAuth 2.0 client secrets file and rename it to "google_client_secret.json"
#    and place it in the same directory as this script.
#
# 4. Find your Google Calendar ID from the calendar's settings page and paste it below.
# =========================================


# =========================================
# CONFIGURATION
# =========================================
# --- Google Calendar Settings ---
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar"]
GOOGLE_TOKEN_FILE = "google_token.json"
GOOGLE_CREDENTIALS_FILE = "google_client_secret.json"

# IMPORTANT: Replace with your target Google Calendar ID
GOOGLE_CALENDAR_ID = "YOUR GOOGLE CALENDAR IS - FROM the SETTINGS SCREEN OF YOUR CALENDAR"

# --- Sync Settings ---
SYNC_LOOKBACK_DAYS = 30
SYNC_LOOKAHEAD_DAYS = 30
SYNC_INTERVAL_MINUTES = 15
LOCAL_TZ = "Australia/Sydney"

# =========================================
# GUI SETUP & THREAD-SAFE LOGGING
# =========================================
log_queue = queue.Queue()
gui_root = tk.Tk()
gui_root.title("Outlook → Google Sync")

# --- Frame for buttons ---
button_frame = tk.Frame(gui_root)
button_frame.pack(pady=5)

start_button = tk.Button(button_frame, text="Start Sync", command=lambda: start_sync_action())
start_button.pack(side=tk.LEFT, padx=5)

stop_button = tk.Button(button_frame, text="Stop Sync", state=tk.DISABLED, command=lambda: stop_sync_action())
stop_button.pack(side=tk.LEFT, padx=5)

# --- Log window ---
gui_log = ScrolledText(gui_root, width=120, height=35, state='disabled')
gui_log.pack(padx=10, pady=10)

def log(msg):
    """Puts a message into the thread-safe queue."""
    now = datetime.datetime.now().strftime("%Y-m-%d %H:%M:%S")
    log_queue.put(f"[{now}] {msg}\n")

def process_log_queue():
    """Processes messages from the queue and updates the GUI."""
    try:
        while not log_queue.empty():
            message = log_queue.get_nowait()
            gui_log.config(state='normal')
            gui_log.insert(tk.END, message)
            gui_log.see(tk.END)
            gui_log.config(state='disabled')
    finally:
        gui_root.after(200, process_log_queue)

# =========================================
# THREADING CONTROL
# =========================================
sync_thread = None
stop_event = threading.Event()

def start_sync_action():
    """Starts the background sync thread."""
    global sync_thread
    if sync_thread and sync_thread.is_alive():
        log("Sync is already running.")
        return

    stop_event.clear()
    sync_thread = threading.Thread(target=background_sync, daemon=True)
    sync_thread.start()
    
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)
    log("Sync process started.")

def stop_sync_action():
    """Stops the background sync thread."""
    if not sync_thread or not sync_thread.is_alive():
        log("Sync is not currently running.")
        return
        
    stop_event.set()
    start_button.config(state=tk.NORMAL)
    stop_button.config(state=tk.DISABLED)
    log("Sync process stopped. Will halt after the current cycle finishes.")

# =========================================
# GOOGLE AUTH & API
# =========================================
def get_google_service():
    """Authenticates with Google and returns a service object."""
    creds = None
    if os.path.exists(GOOGLE_TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(GOOGLE_TOKEN_FILE, GOOGLE_SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                log(f"Error refreshing Google token: {e}")
                creds = None
        
        if not creds:
            if not os.path.exists(GOOGLE_CREDENTIALS_FILE):
                log(f"FATAL: Google credentials file not found at '{GOOGLE_CREDENTIALS_FILE}'")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(GOOGLE_CREDENTIALS_FILE, GOOGLE_SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(GOOGLE_TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    service = build("calendar", "v3", credentials=creds)
    log("Google service initialized successfully.")
    log(f"Syncing to calendar: {GOOGLE_CALENDAR_ID}")
    return service

# =========================================
# OUTLOOK LOCAL CLIENT
# =========================================
def fetch_outlook_events():
    """Fetches calendar events from the local Outlook client."""
    pythoncom.CoInitialize()
    try:
        log("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9)
        log("Outlook calendar accessed.")

        now = datetime.datetime.now()
        start_time = now - datetime.timedelta(days=SYNC_LOOKBACK_DAYS)
        end_time = now + datetime.timedelta(days=SYNC_LOOKAHEAD_DAYS)

        start_str = start_time.strftime('%m/%d/%Y %H:%M %p')
        end_str = end_time.strftime('%m/%d/%Y %H:%M %p')

        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        
        restriction = f"[Start] >= '{start_str}' AND [End] <= '{end_str}'"
        items = items.Restrict(restriction)

        events = []
        count_skipped = 0
        for appt in items:
            if appt.AllDayEvent:
                continue
            
            try:
                events.append({
                    "EntryID": appt.EntryID,
                    "Subject": appt.Subject,
                    "Start": appt.Start,
                    "End": appt.End,
                    "Location": appt.Location,
                    "Body": appt.Body,
                })
            except Exception as e:
                count_skipped += 1
                log(f"Skipping an event due to error: {e}")

        log(f"Fetched {len(events)} Outlook events (Skipped: {count_skipped})")
        return events

    except Exception as e:
        log(f"Failed to connect to Outlook: {e}")
        return []
    finally:
        pythoncom.CoUninitialize()

# =========================================
# SYNC LOGIC
# =========================================
def sync_calendars():
    """Main function to perform the sync between Outlook and Google."""
    try:
        log("Starting sync cycle...")
        outlook_events = fetch_outlook_events()
        
        service = get_google_service()
        if not service:
            log("Could not initialize Google service. Skipping sync.")
            return

        now_utc = datetime.datetime.now(pytz.utc)
        time_min = (now_utc - datetime.timedelta(days=SYNC_LOOKBACK_DAYS)).isoformat()
        time_max = (now_utc + datetime.timedelta(days=SYNC_LOOKAHEAD_DAYS)).isoformat()

        google_events_result = service.events().list(
            calendarId=GOOGLE_CALENDAR_ID,
            timeMin=time_min,
            timeMax=time_max,
            showDeleted=False,
            singleEvents=True,
            maxResults=2500
        ).execute()
        google_events = google_events_result.get("items", [])

        google_event_map = {}
        for g_event in google_events:
            oid = g_event.get("extendedProperties", {}).get("private", {}).get("OutlookEntryID")
            if oid:
                google_event_map[oid] = g_event
        
        log(f"Found {len(google_events)} Google events in sync window.")

        created, updated, deleted = 0, 0, 0
        outlook_ids_in_sync_window = set()
        
        local_tz = pytz.timezone(LOCAL_TZ)

        for o_event in outlook_events:
            oid = o_event["EntryID"]
            outlook_ids_in_sync_window.add(oid)

            start_naive = datetime.datetime(
                o_event["Start"].year, o_event["Start"].month, o_event["Start"].day,
                o_event["Start"].hour, o_event["Start"].minute, o_event["Start"].second
            )
            end_naive = datetime.datetime(
                o_event["End"].year, o_event["End"].month, o_event["End"].day,
                o_event["End"].hour, o_event["End"].minute, o_event["End"].second
            )

            start_local = local_tz.localize(start_naive)
            end_local = local_tz.localize(end_naive)
            
            start_utc = start_local.astimezone(pytz.utc)
            end_utc = end_local.astimezone(pytz.utc)

            event_data = {
                "summary": o_event["Subject"],
                "location": o_event["Location"],
                "description": o_event["Body"],
                "start": {"dateTime": start_utc.isoformat(), "timeZone": "UTC"},
                "end": {"dateTime": end_utc.isoformat(), "timeZone": "UTC"},
                "extendedProperties": {"private": {"OutlookEntryID": oid}},
            }

            g_event = google_event_map.get(oid)

            if g_event:
                g_start = datetime.datetime.fromisoformat(g_event['start']['dateTime'])
                g_end = datetime.datetime.fromisoformat(g_event['end']['dateTime'])
                
                if (g_event.get('summary') != event_data['summary'] or
                    g_event.get('location') != event_data['location'] or
                    g_event.get('description') != event_data['description'] or
                    g_start != start_utc or
                    g_end != end_utc):
                    
                    service.events().update(
                        calendarId=GOOGLE_CALENDAR_ID, 
                        eventId=g_event["id"], 
                        body=event_data
                    ).execute()
                    updated += 1
                    log(f"Updated: {o_event['Subject']}")
            else:
                service.events().insert(
                    calendarId=GOOGLE_CALENDAR_ID, 
                    body=event_data
                ).execute()
                created += 1
                log(f"Created: {o_event['Subject']}")

        for oid, g_event in google_event_map.items():
            if oid not in outlook_ids_in_sync_window:
                try:
                    service.events().delete(
                        calendarId=GOOGLE_CALENDAR_ID, 
                        eventId=g_event["id"]
                    ).execute()
                    deleted += 1
                    log(f"Deleted: {g_event.get('summary')}")
                except Exception as e:
                    log(f"Could not delete event '{g_event.get('summary')}': {e}")

        log(f"Sync completed: {created} created, {updated} updated, {deleted} deleted.")

    except Exception as e:
        log(f"FATAL ERROR during sync: {e}")

# =========================================
# BACKGROUND SYNC THREAD
# =========================================
def background_sync():
    """Wrapper to run the sync logic in a continuous loop until stopped."""
    while not stop_event.is_set():
        sync_calendars()
        log(f"Waiting for {SYNC_INTERVAL_MINUTES} minutes until next sync...")
        # Use wait() on the event object, which can be interrupted by set()
        stop_event.wait(SYNC_INTERVAL_MINUTES * 60)

# =========================================
# MAIN EXECUTION
# =========================================
if __name__ == "__main__":
    log("Starting Outlook to Google Calendar Sync application.")
    log("Press 'Start Sync' to begin.")
    
    process_log_queue()
    gui_root.mainloop()
