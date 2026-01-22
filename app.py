import json
import logging
import os
import queue
import threading
import webbrowser
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from tkinter import BOTH, Button, Canvas, Entry, Frame, Label, Scrollbar, StringVar, Tk, Toplevel
from tkinter import messagebox
from zoneinfo import ZoneInfo

import msal
import requests
from dotenv import load_dotenv

SETTINGS_FILE = Path("settings.json")
ACCOUNTS_FILE = Path("accounts.txt")
LOG_FILE = Path("outlook_clock.log")
EVENT_REFRESH_SECONDS = 300
TIME_REFRESH_MILLISECONDS = 1000
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
LOCAL_TZ = ZoneInfo("America/New_York")
CURRENT_EVENT_FONT = ("Helvetica", 28)
NEXT_EVENT_FONT = ("Helvetica", 14)
NEXT_EVENT_COLOR = "#CCCCCC"
WINDOWS_TZ_MAP = {
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver",
    "Pacific Standard Time": "America/Los_Angeles",
}

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class TenantConfig:
    name: str
    client_id: str
    tenant_id: str
    user_email: str


@dataclass(frozen=True)
class EventDisplay:
    day_label: str
    time_label: str
    subject: str
    start_time: datetime
    end_time: datetime | None


def is_event_active(start_time: datetime | None, end_time: datetime | None) -> bool:
    if not start_time or not end_time:
        return False
    now_local = datetime.now(LOCAL_TZ)
    return start_time.astimezone(LOCAL_TZ) <= now_local <= end_time.astimezone(LOCAL_TZ)


def is_event_soon(start_time: datetime | None, minutes: int = 10) -> bool:
    if not start_time:
        return False
    now_local = datetime.now(LOCAL_TZ)
    start_local = start_time.astimezone(LOCAL_TZ)
    if start_local <= now_local:
        return False
    return start_local - now_local <= timedelta(minutes=minutes)


def format_time_until(start_time: datetime | None) -> str:
    if not start_time:
        return ""
    now_local = datetime.now(LOCAL_TZ)
    start_local = start_time.astimezone(LOCAL_TZ)
    if start_local <= now_local:
        return ""
    delta = start_local - now_local
    total_minutes = int(delta.total_seconds() // 60)
    hours, minutes = divmod(total_minutes, 60)
    if hours and minutes:
        return f"In {hours} hr {minutes} min"
    if hours:
        return f"In {hours} hr"
    return f"In {minutes} min"


def load_settings():
    load_dotenv()
    client_id = os.getenv("CLIENT_ID")
    tenant_id = os.getenv("TENANT_ID") or "common"
    user_email = os.getenv("USER_EMAIL", "")

    if SETTINGS_FILE.exists():
        settings = json.loads(SETTINGS_FILE.read_text())
        tenants = []
        for item in settings.get("tenants", []):
            name = item.get("name") or item.get("tenant_name") or "Tenant"
            tenants.append(
                TenantConfig(
                    name=name,
                    client_id=item["client_id"],
                    tenant_id=item.get("tenant_id") or "common",
                    user_email=item.get("user_email", ""),
                )
            )
        if tenants:
            return tenants

    if not client_id:
        raise ValueError("CLIENT_ID is required in the environment or .env file.")

    return [
        TenantConfig(
            name="Tenant",
            client_id=client_id,
            tenant_id=tenant_id,
            user_email=user_email,
        )
    ]


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


def format_event_time(event_start: datetime, event_end: datetime | None) -> tuple[str, str]:
    local_time = event_start.astimezone(LOCAL_TZ)
    end_time = event_end.astimezone(LOCAL_TZ) if event_end else None
    today = datetime.now(LOCAL_TZ).date()
    tomorrow = today + timedelta(days=1)

    if local_time.date() == today:
        day_label = "Today"
    elif local_time.date() == tomorrow:
        day_label = "Tomorrow"
    else:
        day_label = local_time.strftime("%b %d, %Y")

    if end_time:
        time_label = f"{local_time.strftime('%I:%M %p')} - {end_time.strftime('%I:%M %p %Z')}"
    else:
        time_label = local_time.strftime("%I:%M %p %Z")
    return day_label, time_label


def parse_graph_datetime(value: str, time_zone: str) -> datetime | None:
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        try:
            mapped_zone = WINDOWS_TZ_MAP.get(time_zone, time_zone)
            parsed = parsed.replace(tzinfo=ZoneInfo(mapped_zone))
        except Exception:  # noqa: BLE001 - fallback to UTC
            parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def get_next_events(access_token, user_email, tenant_label, count=2) -> list[EventDisplay]:
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=1)

    params = {
        "startDateTime": now.isoformat(),
        "endDateTime": end.isoformat(),
        "$orderby": "start/dateTime",
        "$top": str(count),
    }

    headers = {"Authorization": f"Bearer {access_token}"}

    url = f"{GRAPH_ENDPOINT}/me/calendarView"
    response = requests.get(url, headers=headers, params=params, timeout=10)
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        detail = response.text.strip() if response.text else str(exc)
        raise RuntimeError(f"{tenant_label}: Graph request failed ({response.status_code}). {detail}") from exc

    data = response.json()
    events = data.get("value", [])
    if not events:
        return []

    displays: list[EventDisplay] = []
    for event in events:
        subject = event.get("subject") or "(No subject)"
        start = event.get("start", {})
        end_info = event.get("end", {})
        start_time = start.get("dateTime", "")
        start_zone = start.get("timeZone", "UTC")
        end_time = end_info.get("dateTime", "")
        end_zone = end_info.get("timeZone", "UTC")
        try:
            parsed = parse_graph_datetime(start_time, start_zone)
            parsed_end = parse_graph_datetime(end_time, end_zone)
            if parsed is None:
                raise ValueError("Invalid start time")
            day_label, time_label = format_event_time(parsed, parsed_end)
        except ValueError:
            continue

        if user_email:
            subject = f"{subject} ({user_email})"

        displays.append(
            EventDisplay(
                day_label=day_label,
                time_label=time_label,
                subject=subject,
                start_time=parsed,
                end_time=parsed_end,
            )
        )

    return displays


def select_next_events(events: list[EventDisplay]) -> tuple[EventDisplay | None, EventDisplay | None]:
    if not events:
        return None, None
    events_sorted = sorted(events, key=lambda event: event.start_time)
    current = events_sorted[0]
    next_event = events_sorted[1] if len(events_sorted) > 1 else None
    return current, next_event


def parse_accounts_file(path: Path) -> list[TenantConfig]:
    if not path.exists():
        logger.warning("Accounts file not found at %s", path)
        return []
    content = path.read_text().splitlines()
    tenants: list[TenantConfig] = []
    current: dict[str, str] = {}
    for line in content:
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip()
        if key.lower() == "tenant_name" and current:
            tenants.append(
                TenantConfig(
                    name=current.get("Tenant_name") or current.get("TENANT_NAME") or "Tenant",
                    client_id=current.get("CLIENT_ID", ""),
                    tenant_id=current.get("TENANT_ID") or "common",
                    user_email=current.get("USER_EMAIL", ""),
                )
            )
            current = {}
        current[key] = value
    if current:
        tenants.append(
            TenantConfig(
                name=current.get("Tenant_name") or current.get("TENANT_NAME") or "Tenant",
                client_id=current.get("CLIENT_ID", ""),
                tenant_id=current.get("TENANT_ID") or "common",
                user_email=current.get("USER_EMAIL", ""),
            )
        )
    return [tenant for tenant in tenants if tenant.client_id]


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

        self.canvas = Canvas(self.window, bg="black", highlightthickness=0)
        self.scrollbar = Scrollbar(self.window, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill=BOTH, expand=True)

        self.container = Frame(self.canvas, bg="black")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.container, anchor="nw")
        self.container.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)

        Button(self.window, text="Import accounts.txt", command=self.import_accounts).pack(pady=5)
        Button(self.window, text="Save", command=self.save).pack(pady=5)

    def on_frame_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)

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
            tenant_name = StringVar()
            client_id = StringVar()
            tenant_id = StringVar(value="common")
            user_email = StringVar()
            Label(self.container, text="Tenant name", fg="white", bg="black").pack(anchor="w", padx=20)
            Entry(self.container, textvariable=tenant_name).pack(fill="x", padx=20)
            Label(self.container, text="Client ID", fg="white", bg="black").pack(anchor="w", padx=20)
            Entry(self.container, textvariable=client_id).pack(fill="x", padx=20)
            Label(self.container, text="Tenant ID", fg="white", bg="black").pack(anchor="w", padx=20)
            Entry(self.container, textvariable=tenant_id).pack(fill="x", padx=20)
            Label(self.container, text="User email", fg="white", bg="black").pack(anchor="w", padx=20)
            Entry(self.container, textvariable=user_email).pack(fill="x", padx=20)
            Button(
                self.container,
                text="Login",
                command=lambda idx=index: self.login_tenant(idx),
            ).pack(anchor="e", padx=20, pady=5)
            self.entries.append((tenant_name, client_id, tenant_id, user_email))

    def save(self):
        tenants = []
        for tenant_name, client_id, tenant_id, user_email in self.entries:
            if client_id.get().strip():
                tenants.append(
                    {
                        "name": tenant_name.get().strip() or "Tenant",
                        "client_id": client_id.get().strip(),
                        "tenant_id": tenant_id.get().strip(),
                        "user_email": user_email.get().strip(),
                    }
                )
        if tenants:
            SETTINGS_FILE.write_text(json.dumps({"tenants": tenants}, indent=2))
        self.window.destroy()

    def import_accounts(self):
        tenants = parse_accounts_file(ACCOUNTS_FILE)
        if not tenants:
            return
        self.count_var.set(str(len(tenants)))
        self.build_entries()
        for entry_vars, tenant in zip(self.entries, tenants):
            tenant_name, client_id, tenant_id, user_email = entry_vars
            tenant_name.set(tenant.name)
            client_id.set(tenant.client_id)
            tenant_id.set(tenant.tenant_id)
            user_email.set(tenant.user_email)

    def login_tenant(self, index):
        try:
            tenant_name, client_id, tenant_id, _user_email = self.entries[index]
        except IndexError:
            return
        client_id_value = client_id.get().strip()
        if not client_id_value:
            messagebox.showerror("Missing client ID", "Please enter a client ID before logging in.")
            return
        tenant_id_value = tenant_id.get().strip() or "common"
        tenant_label = tenant_name.get().strip() or f"Tenant {index + 1}"
        thread = threading.Thread(
            target=self.run_device_flow,
            args=(index, tenant_label, client_id_value, tenant_id_value),
            daemon=True,
        )
        thread.start()

    def run_device_flow(self, index, tenant_label, client_id, tenant_id):
        cache_path = Path(f"token_cache_{index}.json")
        msal_app, cache = build_msal_app(client_id, tenant_id, cache_path)
        flow = msal_app.initiate_device_flow(scopes=["Calendars.Read"])
        if "user_code" not in flow:
            error = flow.get("error")
            description = flow.get("error_description")
            detail = f"{error}: {description}" if error or description else "Unknown error."
            self.root.after(0, lambda: messagebox.showerror("Login failed", detail))
            return
        print(f"{tenant_label} login: {flow['message']}", flush=True)
        webbrowser.open("https://microsoft.com/device")
        self.root.after(
            0,
            lambda: messagebox.showinfo(
                f"Device Login - {tenant_label}",
                f"{tenant_label} login:\n{flow['message']}",
            ),
        )
        token_result = msal_app.acquire_token_by_device_flow(flow)
        save_cache(cache, cache_path)
        if "access_token" not in token_result:
            error = token_result.get("error")
            description = token_result.get("error_description")
            detail = f"{error}: {description}" if error or description else "Unknown error."
            self.root.after(0, lambda: messagebox.showerror("Login failed", detail))
            return
        self.root.after(0, lambda: messagebox.showinfo("Login complete", "Authentication succeeded."))


class OutlookClockApp:
    def __init__(self, root, tenants):
        self.root = root
        self.tenants = tenants
        self.event_queue = queue.Queue()
        self.current_event_day = "Fetching next event..."
        self.current_event_time = ""
        self.current_event_detail = ""
        self.current_event_start: datetime | None = None
        self.current_event_end: datetime | None = None
        self.next_event_day = ""
        self.next_event_time = ""
        self.next_event_detail = ""
        self.next_event_start: datetime | None = None
        self.next_event_end: datetime | None = None

        self.root.title("Outlook Clock")
        self.root.configure(bg="black")
        self.root.attributes("-fullscreen", True)

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=2)
        self.root.grid_rowconfigure(2, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.time_frame = Frame(root, bg="black")
        self.time_frame.grid(row=0, column=0, sticky="nsew")
        self.event_frame = Frame(root, bg="black")
        self.event_frame.grid(row=1, column=0, sticky="nsew")
        self.next_frame = Frame(root, bg="black")
        self.next_frame.grid(row=2, column=0, sticky="nsew")

        self.time_label = Label(
            self.time_frame,
            text="",
            font=("Helvetica", 72),
            fg="white",
            bg="black",
        )
        self.time_label.pack(fill=BOTH, expand=True)

        self.event_label = Label(
            self.event_frame,
            text="",
            font=CURRENT_EVENT_FONT,
            fg="white",
            bg="black",
            wraplength=900,
            justify="center",
            anchor="center",
        )
        self.event_label.pack(fill=BOTH, expand=True)

        self.next_event_label = Label(
            self.next_frame,
            text="",
            font=NEXT_EVENT_FONT,
            fg=NEXT_EVENT_COLOR,
            bg="black",
            wraplength=900,
            justify="center",
            anchor="center",
        )
        self.next_event_label.pack(fill=BOTH, expand=True)

        self.update_time()
        self.schedule_event_refresh()

        root.bind("<Escape>", lambda _event: root.destroy())

    def update_time(self):
        now = datetime.now(LOCAL_TZ)
        self.time_label.config(text=now.strftime("%I:%M:%S %p %Z"))
        if self.current_event_end and now > self.current_event_end.astimezone(LOCAL_TZ):
            if self.next_event_start:
                self.current_event_day = self.next_event_day
                self.current_event_time = self.next_event_time
                self.current_event_detail = self.next_event_detail
                self.current_event_start = self.next_event_start
                self.current_event_end = self.next_event_end
                self.next_event_day = ""
                self.next_event_time = ""
                self.next_event_detail = ""
                self.next_event_start = None
                self.next_event_end = None
            else:
                self.current_event_day = "Fetching next event..."
                self.current_event_time = ""
                self.current_event_detail = ""
                self.current_event_start = None
                self.current_event_end = None
        is_active = is_event_active(self.current_event_start, self.current_event_end)
        is_soon = is_event_soon(self.current_event_start)
        day_label = self.current_event_day
        if self.current_event_day == "Today":
            if is_active:
                day_label = "Today - Currently Occurring"
            else:
                countdown = format_time_until(self.current_event_start)
                if countdown:
                    day_label = f"Today - {countdown}"
        self.event_label.config(
            text=(
                f"{day_label}\n"
                f"{self.current_event_time}\n"
                f"{self.current_event_detail}"
            )
        )
        if is_active:
            event_color = "green"
        elif is_soon:
            event_color = "red"
        else:
            event_color = "white"
        self.event_label.config(fg=event_color)
        soon_next = is_event_soon(self.next_event_start)
        next_day_label = self.next_event_day
        if self.next_event_day == "Today":
            countdown = format_time_until(self.next_event_start)
            if countdown:
                next_day_label = f"Today - {countdown}"
        self.next_event_label.config(
            text=(
                f"{next_day_label}\n"
                f"{self.next_event_time}\n"
                f"{self.next_event_detail}"
            ).strip()
        )
        self.next_event_label.config(fg="red" if soon_next else NEXT_EVENT_COLOR)
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
                events.extend(
                    get_next_events(
                        token,
                        tenant.user_email,
                        tenant.name,
                    )
                )
            current_event, next_event = select_next_events(events)
            if current_event:
                self.event_queue.put(
                    (
                        current_event.day_label,
                        current_event.time_label,
                        current_event.subject,
                        current_event.start_time,
                        current_event.end_time,
                        next_event.start_time if next_event else None,
                        next_event.end_time if next_event else None,
                        next_event.day_label if next_event else "",
                        next_event.time_label if next_event else "",
                        next_event.subject if next_event else "",
                    )
                )
            else:
                self.event_queue.put(("No upcoming events", "", "", None, None, None, None, "", "", ""))
        except Exception as exc:  # noqa: BLE001 - surface errors for display
            logger.exception("Failed to refresh event.")
            self.event_queue.put(("Error", "", str(exc), None, None, None, None, "", "", ""))

    def flush_event_queue(self):
        try:
            while True:
                (
                    day_label,
                    time_info,
                    subject,
                    start_time,
                    end_time,
                    next_start,
                    next_end,
                    next_day,
                    next_time,
                    next_detail,
                ) = self.event_queue.get_nowait()
                self.current_event_day = day_label
                self.current_event_time = time_info
                self.current_event_detail = subject
                self.current_event_start = start_time
                self.current_event_end = end_time
                self.next_event_start = next_start
                self.next_event_end = next_end
                self.next_event_day = next_day
                self.next_event_time = next_time
                self.next_event_detail = next_detail
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
