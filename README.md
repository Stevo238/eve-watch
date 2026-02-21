# Screen Color Tone Watcher

A small Windows desktop app that watches a configurable rectangle on your screen.
If the target color appears in that zone, it plays an audible tone.

## Features
- Editable watch zone (`x`, `y`, `width`, `height`)
- Optional drag-to-select zone overlay
- Live capture preview window to verify what the app can see
- Up to 3 configurable target colors (hex or RGB)
- Adjustable tone frequency and duration
- Configurable color tolerance and scan interval
- One-click temporary silence (`Silence Now`) with configurable duration
- One-click `Test Tone` button for audio verification
- Start/Stop monitoring from a simple UI
- Save/Load profile buttons for settings persistence
- Auto-load profile on startup and auto-save on start/exit (`profile.json`)

## Requirements
- Windows
- Python 3.10+

## Setup
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run
```powershell
python app.py
```

## Notes
- Beep playback uses Windows `winsound.Beep`, so frequency and duration are adjustable.
- Tone volume is controlled by your system volume.
- The app auto-selects a capture backend: `DXcam` first (better for many games), then `mss` fallback.
- If your game is in exclusive fullscreen and detection fails, try borderless-windowed mode.
- Profile settings are saved to `%APPDATA%\\eve-watch\\profile.json` in packaged builds.
