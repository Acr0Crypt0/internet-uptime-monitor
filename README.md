# Internet Uptime Monitor

A lightweight Windows/macOS/Linux desktop app that pings `8.8.8.8` at a configurable interval, logs every result to a local SQLite database, and exports clean, color-coded Excel reports of your daily, weekly, and monthly uptime. Runs quietly in the system tray so it stays out of your way.

## Features

- **Background monitoring** — pings on a configurable interval (default every 5 minutes)
- **Local-only storage** — all ping history lives in a SQLite file next to the script; nothing is sent anywhere
- **Live dashboard** — today / this week / this month uptime percentages, updated in real time
- **Scrolling ping log** — color-coded online/offline events with timestamps
- **Excel export** — generates a styled `.xlsx` report with:
  - Summary sheet with KPI cards
  - Daily breakdown (last 30 days)
  - Weekly breakdown (last 8 weeks)
  - Monthly breakdown (last 12 months)
  - A help/legend sheet explaining every column and color
- **System tray support** — minimizes to a tray icon that changes color based on current connection status
- **No telemetry, no accounts, no cloud** — everything stays on your machine

## Requirements

- Python 3.8+
- `tkinter` (included with most standard Python installs)

Optional, but recommended:

| Package | Enables |
|---|---|
| `pystray` + `Pillow` | System tray icon |
| `openpyxl` | Excel report export |

The app runs without these — it just disables the corresponding feature and shows a reminder in the console.

## Installation

```bash
git clone https://github.com/Acr0Crypt0/internet-uptime-monitor.git
cd internet-uptime-monitor
pip install pystray Pillow openpyxl
```

## Usage

```bash
python internet_monitor.py
```

(On macOS/Linux this may be `python3` depending on your setup.) No administrator privileges are required.

### Optional: hide the console window on Windows

If you'd rather not see a console window at all while the app runs, use `pythonw` instead of `python`, or rename the file to `internet_monitor.pyw` (Windows runs `.pyw` files with the windowless interpreter automatically). This is purely cosmetic — the app behaves identically either way.

## Configuration

Settings are stored in `config.json`, created automatically next to the script on first run:

```json
{
  "interval_minutes": 5,
  "ping_count": 1,
  "ping_timeout": 4,
  "start_minimized": false
}
```

| Key | Description |
|---|---|
| `interval_minutes` | Minutes between each ping check (1–60) |
| `ping_count` | Number of ping packets sent per check |
| `ping_timeout` | Timeout in seconds before a ping is considered failed |
| `start_minimized` | Launch directly to the system tray |

The interval can also be changed live from the app UI without editing the file directly.

## Data

Ping history is stored in `uptime_data.db` (SQLite), created next to the script. Each row records a Unix timestamp and a 1/0 online flag. Delete this file to reset all history.

## Exporting Reports

Click **Export Excel** in the app, or use the **Export Excel…** option from the tray menu. You'll be prompted for a save location; the report covers the full available history at the time of export.

## Building a Standalone Executable (optional)

To distribute the app without requiring Python installed, package it with [PyInstaller](https://pyinstaller.org/):

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile internet_monitor.py
```

The `--noconsole` flag is important — it suppresses the console window the same way `pythonw` does when running from source.

## License

MIT — feel free to use, modify, and distribute.

## Contributing

Issues and pull requests are welcome. If you run into a bug, please include your OS and Python version.
