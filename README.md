# Raspberry Pi Outlook Clock

This project provides a lightweight Tkinter app for a Raspberry Pi 4 that displays the current time and the next upcoming Outlook calendar appointment. It uses the Microsoft Graph API and the device code flow for authentication.

## Features

- Fullscreen clock display optimized for a wall-mounted Raspberry Pi.
- Fetches the next Outlook appointment from Microsoft Graph.
- Token cache persisted locally to avoid repeated sign-ins.

## Prerequisites

- Raspberry Pi OS with Python 3.9+.
- A Microsoft Azure app registration with **Calendars.Read** delegated permissions.

## Setup

1. **Install system packages** (Tkinter is not always installed by default on Raspberry Pi OS):

   ```bash
   sudo apt-get update
   sudo apt-get install -y python3-tk
   ```

2. **Clone this repo and install Python dependencies**:

   ```bash
   git clone <repo-url>
   cd Murder_On_The_Mergviglia
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

3. **Create an Azure app registration**:

   - Go to <https://portal.azure.com> → **Microsoft Entra ID** → **App registrations** → **New registration**.
   - Name it (e.g., `Pi Outlook Clock`).
   - Supported account types: choose what fits your tenant (or use **Accounts in any organizational directory and personal Microsoft accounts**).
   - No redirect URI is required for the device code flow.
   - Under **API permissions**, add **Microsoft Graph** → **Delegated permissions** → **Calendars.Read**.

4. **Configure environment variables**:

   Create a `.env` file in the project root:

   ```ini
   CLIENT_ID=your-client-id-here
   TENANT_ID=common
   USER_EMAIL=you@example.com
   ```

   - `TENANT_ID` can be `common` for multi-tenant, or your tenant GUID.

5. **Run the app**:

   ```bash
   python app.py
   ```

## Autostart on boot (optional)

Create a systemd unit to start the app on boot:

```ini
# /etc/systemd/system/outlook-clock.service
[Unit]
Description=Outlook Clock
After=network-online.target
Wants=network-online.target

[Service]
User=pi
WorkingDirectory=/home/pi/Murder_On_The_Mergviglia
Environment=DISPLAY=:0
ExecStart=/home/pi/Murder_On_The_Mergviglia/.venv/bin/python /home/pi/Murder_On_The_Mergviglia/app.py
Restart=always

[Install]
WantedBy=graphical.target
```

Then enable it:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now outlook-clock.service
```

## Troubleshooting

- If the app can’t open the GUI, confirm the Pi is running a graphical session and that the `DISPLAY` variable is correct.
- Delete `token_cache.json` if you need to force a new login.
- If you see `ModuleNotFoundError: No module named 'msal'`, make sure the virtual environment is active and dependencies are installed:

  ```bash
  source .venv/bin/activate
  pip install -r requirements.txt
  ```
- If you start the app from the Pi desktop and still see `No module named 'msal'`, double-check that you activated the same `.venv` in that desktop terminal before running `python app.py`.
- If you see `ValueError: CLIENT_ID is required in the environment or .env file.`, create the `.env` file in the project root and add `CLIENT_ID`, `TENANT_ID`, and `USER_EMAIL` as shown in the setup steps.
- If you see `no display name and no $DISPLAY environment variable`, you are running without a graphical session. Run the app from a desktop session or export a valid display (for example `export DISPLAY=:0`) and ensure you have an X server running.
- If you see `Failed to start device code flow`, confirm that:
  - The `CLIENT_ID` and `TENANT_ID` values are correct.
  - The Azure app registration is configured as a **public client** (Allow public client flows = **Yes**).
  - The Pi has network access and correct system time (device flow can fail with clock drift). Run:
    ```bash
    date
    timedatectl status
    ```
- If the error persists, check `outlook_clock.log` in the project root for the detailed error information.
- If the log shows `AADSTS50059: No tenant-identifying information found`, verify that `TENANT_ID` is set (or remove it to use the default `common` value).
- If you see `AADSTS7000218` (client secret required), enable **Allow public client flows** in the app registration or switch to a client-secret-based flow.
