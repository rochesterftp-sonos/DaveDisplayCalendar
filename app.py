import json
import logging
import os
import queue
import threading
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tkinter import BOTH, Button, Entry, Frame, Label, StringVar, Tk, Toplevel
from zoneinfo import ZoneInfo

import msal
import requests
from dotenv import load_dotenv

SETTINGS_FILE = Path("settings.json")
LOG_FILE = Path("outlook_clock.log")
EVENT_REFRESH_SECONDS = 300
TIME_REFRESH_MILLISECONDS = 1000
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
LOCAL_TZ = ZoneInfo("America/New_York")

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class TenantConfig:
    client_id: str
    tenant_id: str
    user_email: str


def load_settings():
    load_dotenv()
    client_id = os.getenv("CLIENT_ID")
    tenant_id = os.getenv("TENANT_ID") or "common"
    user_email = os.getenv("USER_EMAIL", "")

    if SETTINGS_FILE.exists():
        settings = json.loads(SETTINGS_FILE.read_text())
        tenants = [
            TenantConfig(
                client_id=item["client_id"],
                tenant_id=item.get("tenant_id") or "common",
                user_email=item.get("user_email", ""),
            )
            for item in settings.get("tenants", [])
        ]
        if tenants:
            return tenants

    if not client_id:
        raise ValueError("CLIENT_ID is required in the environment or .env file.")

    return [TenantConfig(client_id=client_id, tenant_id=tenant_id, user_email=user_email)]


def build_msal_app(client_id, tenant_id, cache_path):
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    logger.info("Using MSAL authority %s", authority)
    cache = msal.SerializableTokenCache()

    if cache_path.exists():
        cache.deserialize(cache_path.read_text())

    app = msal.PublicClientApplication(client_id=client_id, authority=authority, token_cache=cache)
    return app, cache


def save_cache(cache, cache_path):
    if cache.has_state_changed:
        cache_path.write_text(cache.serialize())


def get_access_token(msal_app, cache, cache_path):
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

    save_cache(cache, cache_path)

    if "access_token" not in token_result:
        error = token_result.get("error")
        description = token_result.get("error_description")
        detail = f" ({error}: {description})" if error or description else ""
        if description and "AADSTS7000218" in description:
            hint = (
                " Enable 'Allow public client flows' for the app registration or use "
                "CLIENT_SECRET-based auth."
            )
            detail = f"{detail}{hint}"
        logger.error("Token acquisition failed.%s", detail)
        raise RuntimeError(f"Unable to acquire token.{detail}")

    return token_result["access_token"]


def format_event_time(event_start: datetime) -> tuple[str, str]:
    local_time = event_start.astimezone(LOCAL_TZ)
    today = datetime.now(LOCAL_TZ).date()
    tomorrow = today + timedelta(days=1)

    if local_time.date() == today:
        day_label = "Today"
    elif local_time.date() == tomorrow:
        day_label = "Tomorrow"
    else:
        day_label = local_time.strftime("%b %d, %Y")

    time_label = local_time.strftime("%I:%M %p %Z")
    return day_label, time_label


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
        return "No upcoming events", "", "", None

    event = events[0]
    subject = event.get("subject") or "(No subject)"
    start_time = event.get("start", {}).get("dateTime", "")
    parsed = None
    try:
        parsed = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
        day_label, time_label = format_event_time(parsed)
    except ValueError:
        day_label = "Upcoming event"
        time_label = start_time

    if user_email:
        subject = f"{subject} ({user_email})"

    time_line = time_label
    return day_label, time_line, subject, parsed


def select_next_event(events):
    dated_events = [event for event in events if event[3] is not None]
    if not dated_events:
        if not events:
            return "No upcoming events", "", ""
        day_label, time_label, subject, _event_time = events[0]
        return day_label, time_label, subject
    dated_events.sort(key=lambda item: item[3])
    day_label, time_label, subject, _event_time = dated_events[0]
    return day_label, time_label, subject


class SettingsWindow:
    def __init__(self, root):
        self.root = root
        self.count_var = StringVar(value="1")
        self.entries = []
        self.window = Toplevel(root)
        self.window.title("Outlook Clock Settings")
        self.window.configure(bg="black")
        self.window.geometry("800x600")

        Label(self.window, text="Number of tenants", fg="white", bg="black").pack()
        Entry(self.window, textvariable=self.count_var).pack()
        Button(self.window, text="Set", command=self.build_entries).pack(pady=10)
        self.container = Frame(self.window, bg="black")
        self.container.pack(fill=BOTH, expand=True)
        Button(self.window, text="Save", command=self.save).pack(pady=10)

    def build_entries(self):
        for widget in self.container.winfo_children():
            widget.destroy()
        self.entries.clear()
        try:
            count = max(1, int(self.count_var.get()))
        except ValueError:
            count = 1
        for index in range(count):
            Label(
                self.container,
                text=f"Tenant {index + 1}",
                fg="white",
                bg="black",
            ).pack(anchor="w", padx=10, pady=5)
            client_id = StringVar()
            tenant_id = StringVar(value="common")
            user_email = StringVar()
            Label(self.container, text="Client ID", fg="white", bg="black").pack(anchor="w", padx=20)
            Entry(self.container, textvariable=client_id).pack(fill="x", padx=20)
            Label(self.container, text="Tenant ID", fg="white", bg="black").pack(anchor="w", padx=20)
            Entry(self.container, textvariable=tenant_id).pack(fill="x", padx=20)
            Label(self.container, text="User email", fg="white", bg="black").pack(anchor="w", padx=20)
            Entry(self.container, textvariable=user_email).pack(fill="x", padx=20)
            self.entries.append((client_id, tenant_id, user_email))

    def save(self):
        tenants = []
        for client_id, tenant_id, user_email in self.entries:
            if client_id.get().strip():
                tenants.append(
                    {
                        "client_id": client_id.get().strip(),
                        "tenant_id": tenant_id.get().strip(),
                        "user_email": user_email.get().strip(),
                    }
                )
        if tenants:
            SETTINGS_FILE.write_text(json.dumps({"tenants": tenants}, indent=2))
        self.window.destroy()


class OutlookClockApp:
    def __init__(self, root, tenants):
        self.root = root
        self.tenants = tenants
        self.event_queue = queue.Queue()
        self.current_event_day = "Fetching next event..."
        self.current_event_time = ""
        self.current_event_detail = ""

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
        now = datetime.now(LOCAL_TZ)
        self.time_label.config(text=now.strftime("%I:%M:%S %p %Z"))
        self.event_label.config(
            text=(
                f"{self.current_event_day}\n"
                f"{self.current_event_time}\n"
                f"{self.current_event_detail}"
            )
        )
        self.root.after(TIME_REFRESH_MILLISECONDS, self.update_time)
        self.flush_event_queue()

    def schedule_event_refresh(self):
        thread = threading.Thread(target=self.refresh_event, daemon=True)
        thread.start()
        self.root.after(EVENT_REFRESH_SECONDS * 1000, self.schedule_event_refresh)

    def refresh_event(self):
        try:
            events = []
            for index, tenant in enumerate(self.tenants):
                cache_path = Path(f"token_cache_{index}.json")
                msal_app, cache = build_msal_app(
                    tenant.client_id,
                    tenant.tenant_id,
                    cache_path,
                )
                token = get_access_token(msal_app, cache, cache_path)
                day_label, time_info, subject, event_time = get_next_event(
                    token,
                    tenant.user_email,
                )
                events.append((day_label, time_info, subject, event_time))
            day_label, time_info, subject = select_next_event(events)
            self.event_queue.put((day_label, time_info, subject))
        except Exception as exc:  # noqa: BLE001 - surface errors for display
            logger.exception("Failed to refresh event.")
            self.event_queue.put(("Error", "", str(exc)))

    def flush_event_queue(self):
        try:
            while True:
                day_label, time_info, subject = self.event_queue.get_nowait()
                self.current_event_day = day_label
                self.current_event_time = time_info
                self.current_event_detail = subject
        except queue.Empty:
            return


def main():
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    root = Tk()
    if not SETTINGS_FILE.exists():
        root.withdraw()
        settings_window = SettingsWindow(root)
        root.wait_window(settings_window.window)
        root.deiconify()
    tenants = load_settings()
    logger.info("Loaded settings with %s tenants", len(tenants))
    app = OutlookClockApp(root, tenants)
    root.mainloop()


if __name__ == "__main__":
    main()
