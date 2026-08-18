# Screen Color Tone Watcher

A small desktop app (Windows + Linux) that watches a configurable rectangle on your screen.
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

## Corp Intel reporting

The app can report to the Intel board on the corp site so mates can see who has
eyes where. Any alert — any colour, in any zone — is the flag; the board never
shows which colour tripped, only that this post is in contact.

Fill in the **Corp Intel** section:

| Field | What it is |
|-------|------------|
| Server | The corp site, e.g. `https://eve.motronimeadows.com` |
| Token | Issued on the site's Intel tab under **Watch posts**. Shown once — copy it straight away |
| Pilot | Your character name, shown on the board |
| System | The system you have eyes in. **Required** — no system, no card on the board |
| Post label | Optional, e.g. "Tama gate cam" |

Tick **Report to corp Intel board**, then press Start as usual. The status line
under the section shows whether reporting is working.

Reporting runs on its own thread, so a slow or offline server can never delay
your own audible alert. A heartbeat goes out every 20 seconds carrying the full
state, so a single dropped message corrects itself on the next beat. If the app
stops sending, your card goes stale on the board after 90 seconds rather than
sitting there falsely green.

Settings (token included) are saved in `profile.json` alongside everything else.

## Requirements
- Windows or Linux
- Python 3.10+ (only if running from source; the packaged binaries need nothing)

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

## Linux

Download `eve-watch-linux` from the [Releases page](https://github.com/Stevo238/eve-watch/releases), then:

```bash
chmod +x eve-watch-linux
./eve-watch-linux
```

Or run from source:

```bash
sudo apt install python3-tk
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

**Important — Wayland vs X11:** screen capture uses `mss`, which reads the X11
display. Ubuntu 25.10's default GNOME desktop is Wayland-only, so:

- Apps running through **XWayland** (anything under Wine/Proton — including EVE
  Online via Steam or the official launcher) **can** be captured.
- Native Wayland windows appear **black** to the capture. Use the built-in
  **Preview** button to confirm the app can actually see your target window.
- If capture fails entirely, log into an X11/Xorg session (may require
  installing one on 25.10) or run the game windowed/borderless.

Tones play through `paplay`/`pw-play`/`aplay` (PipeWire and PulseAudio ship with
Ubuntu, so no extra install needed). The profile is saved to
`~/.config/eve-watch/profile.json`.

## Notes
- Windows beep playback uses `winsound`; Linux writes the same generated WAV through PipeWire/ALSA.
- Alert volume is controlled by the in-app volume slider.
- The app auto-selects a capture backend: `DXcam` first on Windows (better for many games), then `mss` fallback (always used on Linux).
- If your game is in exclusive fullscreen and detection fails, try borderless-windowed mode.
- Profile settings are saved to `%APPDATA%\eve-watch\profile.json` (Windows) or `~/.config/eve-watch/profile.json` (Linux) in packaged builds.
