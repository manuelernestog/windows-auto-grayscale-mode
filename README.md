# Windows-Auto-Grayscale-Mode

A Windows application to automatically toggle the grayscale filter based on a schedule or manually.

## Description

Windows-Grayscale-Mode is a utility application for Windows that allows users to control the system's grayscale color filter. It can be toggled manually via a button in the application window or through the system tray icon. Additionally, it supports automatic scheduling, enabling the grayscale filter during specified night hours and disabling it during the day. The application runs in the system tray and can be configured to start automatically with Windows.

## Installation

This project requires Python. You can install the necessary dependencies using pip:

```bash
pip install pystray Pillow pyautogui pyinstaller
```
(Note: `tkinter` and `winreg` are typically included with Python on Windows.)

## Usage

To run the application, execute the main Python script:

```bash
python grayscale_mode_app.py
```

The application will open a window and place an icon in the system tray.

- **Manual Control:** Use the "Toggle Grayscale" button in the main window or the "Toggle Grayscale" option in the tray icon menu to manually activate or deactivate the grayscale filter.
- **Automatic Scheduling:** Check the "Enable Scheduling" box and set the desired "Start Time (HH:MM)" and "End Time (HH:MM)". The application will automatically toggle the grayscale filter according to this schedule.
- **Start with Windows:** Use the tray icon menu options "Start with Windows" or "Do not start with Windows" to configure the application to launch automatically when you log in.

## Configuration

The automatic scheduling settings are stored in the `schedule_config.json` file in the application's configuration directory within your user's AppData folder. The specific path is typically:

`%APPDATA%\RestMode\schedule_config.json`

This file contains the following keys:

- `"start_time"`: String in "HH:MM" format (24-hour clock) for when grayscale should be activated.
- `"end_time"`: String in "HH:MM" format (24-hour clock) for when grayscale should be deactivated.
- `"enabled"`: Boolean (`true` or `false`) to enable or disable automatic scheduling.

Example `schedule_config.json`:
```json
{
    "start_time": "22:00",
    "end_time": "06:00",
    "enabled": true
}
```

## Building

You can create a standalone executable of the application using PyInstaller and the provided spec file:

```bash
pyinstaller grayscale_mode_app.spec
```

The executable will be found in the `dist` directory.
