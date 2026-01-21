import json
import logging
import os
import queue
import threading
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tkinter import BOTH, Label, Tk

import msal
import requests
from dotenv import load_dotenv

TOKEN_CACHE_FILE = Path("token_cache.json")
LOG_FILE = Path("outlook_clock.log")
EVENT_REFRESH_SECONDS = 300
TIME_REFRESH_MILLISECONDS = 1000
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"

logger = logging.getLogger(__name__)


def load_settings():
    load_dotenv()
    client_id = os.getenv("CLIENT_ID")
    tenant_id = os.getenv("TENANT_ID", "common")
    user_email = os.getenv("USER_EMAIL", "")

    if not client_id:
        raise ValueError("CLIENT_ID is required in the environment or .env file.")

    return client_id, tenant_id, user_email


def build_msal_app(client_id, tenant_id):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    cache = msal.SerializableTokenCache()

    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())

    app = msal.PublicClientApplication(client_id=client_id, authority=authority, token_cache=cache)
    return app, cache


def save_cache(cache):
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize())


def get_access_token(msal_app, cache):
    accounts = msal_app.get_accounts()
    scopes = ["Calendars.Read"]

    if accounts:
        token_result = msal_app.acquire_token_silent(scopes, account=accounts[0])
    else:
        token_result = None

    if not token_result:
        flow = msal_app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            error = flow.get("error")
            description = flow.get("error_description")
            detail = f" ({error}: {description})" if error or description else ""
            logger.error("Device code flow failed to start.%s", detail)
            raise RuntimeError(f"Failed to start device code flow.{detail}")

        print(flow["message"], flush=True)
        token_result = msal_app.acquire_token_by_device_flow(flow)

    save_cache(cache)

    if "access_token" not in token_result:
        error = token_result.get("error")
        description = token_result.get("error_description")
        detail = f" ({error}: {description})" if error or description else ""
        logger.error("Token acquisition failed.%s", detail)
        raise RuntimeError(f"Unable to acquire token.{detail}")

    return token_result["access_token"]


def get_next_event(access_token, user_email):
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=1)

    params = {
        "startDateTime": now.isoformat(),
        "endDateTime": end.isoformat(),
        "$orderby": "start/dateTime",
        "$top": "1",
    }

    headers = {"Authorization": f"Bearer {access_token}"}

    url = f"{GRAPH_ENDPOINT}/me/calendarView"
    response = requests.get(url, headers=headers, params=params, timeout=10)
    response.raise_for_status()

    data = response.json()
    events = data.get("value", [])
    if not events:
        return "No upcoming events", ""

    event = events[0]
    subject = event.get("subject") or "(No subject)"
    start_time = event.get("start", {}).get("dateTime", "")
    time_zone = event.get("start", {}).get("timeZone", "UTC")

    try:
        parsed = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
        local_time = parsed.astimezone()
        time_str = local_time.strftime("%I:%M %p")
    except ValueError:
        time_str = start_time

    if user_email:
        subject = f"{subject} ({user_email})"

    return subject, f"{time_str} {time_zone}"


class OutlookClockApp:
    def __init__(self, root, msal_app, cache, user_email):
        self.root = root
        self.msal_app = msal_app
        self.cache = cache
        self.user_email = user_email
        self.event_queue = queue.Queue()
        self.current_event = "Fetching next event..."
        self.current_event_time = ""

        self.root.title("Outlook Clock")
        self.root.configure(bg="black")
        self.root.attributes("-fullscreen", True)

        self.time_label = Label(
            root,
            text="",
            font=("Helvetica", 72),
            fg="white",
            bg="black",
        )
        self.time_label.pack(fill=BOTH, expand=True)

        self.event_label = Label(
            root,
            text="",
            font=("Helvetica", 28),
            fg="white",
            bg="black",
            wraplength=900,
        )
        self.event_label.pack(fill=BOTH, expand=True)

        self.update_time()
        self.schedule_event_refresh()

        root.bind("<Escape>", lambda _event: root.destroy())

    def update_time(self):
        now = datetime.now()
        self.time_label.config(text=now.strftime("%I:%M:%S %p"))
        self.event_label.config(text=f"{self.current_event}\n{self.current_event_time}")
        self.root.after(TIME_REFRESH_MILLISECONDS, self.update_time)
        self.flush_event_queue()

    def schedule_event_refresh(self):
        thread = threading.Thread(target=self.refresh_event, daemon=True)
        thread.start()
        self.root.after(EVENT_REFRESH_SECONDS * 1000, self.schedule_event_refresh)

    def refresh_event(self):
        try:
            token = get_access_token(self.msal_app, self.cache)
            subject, time_info = get_next_event(token, self.user_email)
            self.event_queue.put((subject, time_info))
        except Exception as exc:  # noqa: BLE001 - surface errors for display
            logger.exception("Failed to refresh event.")
            self.event_queue.put((f"Error: {exc}", ""))

    def flush_event_queue(self):
        try:
            while True:
                subject, time_info = self.event_queue.get_nowait()
                self.current_event = subject
                self.current_event_time = time_info
        except queue.Empty:
            return


def main():
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    client_id, tenant_id, user_email = load_settings()
    msal_app, cache = build_msal_app(client_id, tenant_id)

    root = Tk()
    app = OutlookClockApp(root, msal_app, cache, user_email)
    root.mainloop()


if __name__ == "__main__":
    main()
