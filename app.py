import json
import os
import re
import sys
import threading
import time
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox

import winsound
from PIL import Image, ImageTk

try:
    import mss
except Exception:
    mss = None

try:
    import dxcam
except Exception:
    dxcam = None

try:
    import numpy as np
except Exception:
    np = None


@dataclass
class WatchZone:
    x: int
    y: int
    width: int
    height: int


def hex_to_rgb(value: str) -> tuple[int, int, int]:
    cleaned = value.strip()
    if not re.fullmatch(r"#?[0-9a-fA-F]{6}", cleaned):
        raise ValueError("Hex color must be like #FFAA00")
    cleaned = cleaned.lstrip("#")
    return tuple(int(cleaned[i : i + 2], 16) for i in (0, 2, 4))


def bgra_buffer_best_match(
    raw: bytes, targets: list[tuple[int, int, int]], tolerance: int
) -> tuple[int, int]:
    best_count = 0
    best_idx = -1

    for target_idx, (tr, tg, tb) in enumerate(targets):
        count = 0
        for i in range(0, len(raw), 4):
            b = raw[i]
            g = raw[i + 1]
            r = raw[i + 2]
            if abs(r - tr) <= tolerance and abs(g - tg) <= tolerance and abs(b - tb) <= tolerance:
                count += 1
        if count > best_count:
            best_count = count
            best_idx = target_idx

    return best_count, best_idx


def rgb_frame_best_match(frame, targets: list[tuple[int, int, int]], tolerance: int) -> tuple[int, int]:
    if np is None:
        return 0, -1
    if frame is None or frame.size == 0:
        return 0, -1
    if frame.ndim != 3 or frame.shape[2] < 3:
        return 0, -1

    frame_3 = frame[:, :, :3].astype(np.int16)

    best_count = 0
    best_idx = -1

    for target_idx, target in enumerate(targets):
        target_rgb = np.array(target, dtype=np.int16)
        rgb_mask = np.all(np.abs(frame_3 - target_rgb) <= tolerance, axis=2)

        target_bgr = np.array((target[2], target[1], target[0]), dtype=np.int16)
        bgr_mask = np.all(np.abs(frame_3 - target_bgr) <= tolerance, axis=2)

        rgb_count = int(np.count_nonzero(rgb_mask))
        bgr_count = int(np.count_nonzero(bgr_mask))
        count = max(rgb_count, bgr_count)
        if count > best_count:
            best_count = count
            best_idx = target_idx

    return best_count, best_idx


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Screen Color Tone Watcher v1")
        self.root.geometry("560x820")
        self.root.minsize(520, 720)
        self.root.resizable(True, True)
        self.config_path = self._resolve_profile_path()

        self.running = False
        self.monitor_thread: threading.Thread | None = None
        self.capture_backend = "mss"
        self.dx_camera = None
        self.mute_until_ms = 0.0
        self.preview_window: tk.Toplevel | None = None
        self.preview_label: tk.Label | None = None
        self.preview_info: tk.StringVar | None = None
        self.preview_after_id: str | None = None

        self.zone_x = tk.StringVar(value="0")
        self.zone_y = tk.StringVar(value="0")
        self.zone_w = tk.StringVar(value="300")
        self.zone_h = tk.StringVar(value="200")

        self.color_hex_vars = [
            tk.StringVar(value="#7F3107"),
            tk.StringVar(value=""),
            tk.StringVar(value=""),
        ]
        self.color_r_vars = [
            tk.StringVar(value="127"),
            tk.StringVar(value=""),
            tk.StringVar(value=""),
        ]
        self.color_g_vars = [
            tk.StringVar(value="49"),
            tk.StringVar(value=""),
            tk.StringVar(value=""),
        ]
        self.color_b_vars = [
            tk.StringVar(value="7"),
            tk.StringVar(value=""),
            tk.StringVar(value=""),
        ]

        self.tolerance = tk.StringVar(value="15")
        self.interval_ms = tk.StringVar(value="120")
        self.beep_freq = tk.StringVar(value="1200")
        self.beep_dur = tk.StringVar(value="180")
        self.cooldown_ms = tk.StringVar(value="800")
        self.silence_ms = tk.StringVar(value="60000")

        self.status_text = tk.StringVar(value="Status: Idle")

        self._build_ui()
        self.load_profile(show_message=False)

    def _resolve_profile_path(self) -> Path:
        appdata = os.getenv("APPDATA")
        base_dir = Path(appdata) if appdata else (Path.home() / ".config")
        config_dir = base_dir / "eve-watch"
        config_path = config_dir / "profile.json"

        legacy_path = Path(__file__).resolve().parent / "profile.json"
        if not getattr(sys, "frozen", False) and not config_path.exists() and legacy_path.exists():
            return legacy_path

        return config_path

    def _build_ui(self):
        frame = tk.Frame(self.root, padx=12, pady=12)
        frame.pack(fill="both", expand=True)

        zone_box = tk.LabelFrame(frame, text="Watch Zone", padx=10, pady=8)
        zone_box.pack(fill="x", pady=(0, 10))

        self._row(zone_box, "X", self.zone_x, 0)
        self._row(zone_box, "Y", self.zone_y, 1)
        self._row(zone_box, "Width", self.zone_w, 2)
        self._row(zone_box, "Height", self.zone_h, 3)

        tk.Button(zone_box, text="Select Zone on Screen", command=self.select_zone_overlay).grid(
            row=4, column=0, columnspan=2, sticky="we", pady=(8, 0)
        )
        tk.Button(zone_box, text="Preview Capture", command=self.open_preview_window).grid(
            row=5, column=0, columnspan=2, sticky="we", pady=(6, 0)
        )

        color_box = tk.LabelFrame(frame, text="Target Colors (up to 3)", padx=10, pady=8)
        color_box.pack(fill="x", pady=(0, 10))

        tk.Label(color_box, text="Color").grid(row=0, column=0, sticky="w")
        tk.Label(color_box, text="Hex").grid(row=0, column=1, sticky="w")
        tk.Label(color_box, text="R").grid(row=0, column=2, sticky="w")
        tk.Label(color_box, text="G").grid(row=0, column=3, sticky="w")
        tk.Label(color_box, text="B").grid(row=0, column=4, sticky="w")

        for i in range(3):
            row = i * 2 + 1
            tk.Label(color_box, text=f"{i + 1}").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
            tk.Entry(color_box, textvariable=self.color_hex_vars[i], width=12).grid(row=row, column=1, sticky="we", pady=2)
            tk.Entry(color_box, textvariable=self.color_r_vars[i], width=5).grid(row=row, column=2, sticky="we", pady=2)
            tk.Entry(color_box, textvariable=self.color_g_vars[i], width=5).grid(row=row, column=3, sticky="we", pady=2)
            tk.Entry(color_box, textvariable=self.color_b_vars[i], width=5).grid(row=row, column=4, sticky="we", pady=2)

            tk.Button(color_box, text="Hex -> RGB", command=lambda idx=i: self.apply_hex_to_rgb(idx)).grid(
                row=row + 1, column=1, columnspan=2, sticky="we", pady=(0, 4)
            )
            tk.Button(color_box, text="RGB -> Hex", command=lambda idx=i: self.apply_rgb_to_hex(idx)).grid(
                row=row + 1, column=3, columnspan=2, sticky="we", pady=(0, 4)
            )

        for col in range(1, 5):
            color_box.grid_columnconfigure(col, weight=1)

        settings_box = tk.LabelFrame(frame, text="Detection & Tone", padx=10, pady=8)
        settings_box.pack(fill="x", pady=(0, 10))

        self._row(settings_box, "Tolerance (0-255)", self.tolerance, 0)
        self._row(settings_box, "Scan interval ms", self.interval_ms, 1)
        self._row(settings_box, "Beep freq Hz", self.beep_freq, 2)
        self._row(settings_box, "Beep duration ms", self.beep_dur, 3)
        self._row(settings_box, "Cooldown ms", self.cooldown_ms, 4)
        self._row(settings_box, "Silence ms", self.silence_ms, 5)

        controls = tk.Frame(frame)
        controls.pack(fill="x")
        tk.Button(controls, text="Start", command=self.start).pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(controls, text="Stop", command=self.stop).pack(side="left", fill="x", expand=True, padx=(5, 0))

        mute_controls = tk.Frame(frame)
        mute_controls.pack(fill="x", pady=(8, 0))
        tk.Button(mute_controls, text="Silence Now", command=self.silence_for_period).pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        tk.Button(mute_controls, text="Test Tone", command=self.test_tone).pack(
            side="left", fill="x", expand=True, padx=(5, 0)
        )

        profile_controls = tk.Frame(frame)
        profile_controls.pack(fill="x", pady=(8, 0))
        tk.Button(profile_controls, text="Save Profile", command=self.save_profile).pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        tk.Button(profile_controls, text="Load Profile", command=self.load_profile).pack(
            side="left", fill="x", expand=True, padx=(5, 0)
        )

        tk.Label(frame, textvariable=self.status_text, anchor="w").pack(fill="x", pady=(10, 0))

    def _row(self, parent: tk.Widget, label: str, var: tk.StringVar, row: int):
        tk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
        tk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="we", pady=2)
        parent.grid_columnconfigure(1, weight=1)

    def apply_hex_to_rgb(self, idx: int = 0):
        try:
            hex_value = self.color_hex_vars[idx].get().strip()
            if not hex_value:
                messagebox.showerror("Invalid hex", "Hex value cannot be empty")
                return
            r, g, b = hex_to_rgb(hex_value)
        except ValueError as exc:
            messagebox.showerror("Invalid hex", str(exc))
            return
        self.color_r_vars[idx].set(str(r))
        self.color_g_vars[idx].set(str(g))
        self.color_b_vars[idx].set(str(b))

    def apply_rgb_to_hex(self, idx: int = 0):
        try:
            r = int(self.color_r_vars[idx].get())
            g = int(self.color_g_vars[idx].get())
            b = int(self.color_b_vars[idx].get())
        except ValueError:
            messagebox.showerror("Invalid RGB", "R, G, B must be numbers")
            return
        for value in (r, g, b):
            if not 0 <= value <= 255:
                messagebox.showerror("Invalid RGB", "R, G, B must be in 0..255")
                return
        self.color_hex_vars[idx].set(f"#{r:02X}{g:02X}{b:02X}")

    def _parse_color_slot(self, idx: int) -> tuple[int, int, int] | None:
        hex_value = self.color_hex_vars[idx].get().strip()
        r_str = self.color_r_vars[idx].get().strip()
        g_str = self.color_g_vars[idx].get().strip()
        b_str = self.color_b_vars[idx].get().strip()

        if hex_value:
            return hex_to_rgb(hex_value)

        if not (r_str or g_str or b_str):
            return None

        if not (r_str and g_str and b_str):
            raise ValueError(f"Color {idx + 1}: provide all R/G/B values or leave blank")

        r = int(r_str)
        g = int(g_str)
        b = int(b_str)
        for val in (r, g, b):
            if not 0 <= val <= 255:
                raise ValueError(f"Color {idx + 1}: R, G, B must be in 0..255")
        return (r, g, b)

    def _get_profile_dict(self) -> dict:
        return {
            "zone": {
                "x": self.zone_x.get(),
                "y": self.zone_y.get(),
                "width": self.zone_w.get(),
                "height": self.zone_h.get(),
            },
            "color": {
                "colors": [
                    {
                        "hex": self.color_hex_vars[i].get(),
                        "r": self.color_r_vars[i].get(),
                        "g": self.color_g_vars[i].get(),
                        "b": self.color_b_vars[i].get(),
                    }
                    for i in range(3)
                ]
            },
            "detection": {
                "tolerance": self.tolerance.get(),
                "interval_ms": self.interval_ms.get(),
                "cooldown_ms": self.cooldown_ms.get(),
                "silence_ms": self.silence_ms.get(),
            },
            "tone": {
                "beep_freq": self.beep_freq.get(),
                "beep_dur": self.beep_dur.get(),
            },
        }

    def save_profile(self, show_message: bool = True):
        try:
            self.parse_zone()
            _ = self.parse_targets()
            _ = int(self.tolerance.get())
            _ = int(self.interval_ms.get())
            _ = int(self.cooldown_ms.get())
            _ = int(self.silence_ms.get())
            _ = int(self.beep_freq.get())
            _ = int(self.beep_dur.get())

            profile = self._get_profile_dict()
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            self.config_path.write_text(json.dumps(profile, indent=2), encoding="utf-8")
            if show_message:
                messagebox.showinfo("Profile saved", f"Saved to {self.config_path.name}")
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))

    def load_profile(self, show_message: bool = True):
        try:
            if not self.config_path.exists():
                if show_message:
                    messagebox.showinfo("No profile", f"{self.config_path.name} not found yet")
                return

            data = json.loads(self.config_path.read_text(encoding="utf-8"))

            zone = data.get("zone", {})
            self.zone_x.set(str(zone.get("x", self.zone_x.get())))
            self.zone_y.set(str(zone.get("y", self.zone_y.get())))
            self.zone_w.set(str(zone.get("width", self.zone_w.get())))
            self.zone_h.set(str(zone.get("height", self.zone_h.get())))

            color = data.get("color", {})
            colors = color.get("colors")
            if isinstance(colors, list):
                for i in range(3):
                    item = colors[i] if i < len(colors) and isinstance(colors[i], dict) else {}
                    self.color_hex_vars[i].set(str(item.get("hex", self.color_hex_vars[i].get())))
                    self.color_r_vars[i].set(str(item.get("r", self.color_r_vars[i].get())))
                    self.color_g_vars[i].set(str(item.get("g", self.color_g_vars[i].get())))
                    self.color_b_vars[i].set(str(item.get("b", self.color_b_vars[i].get())))
            detection = data.get("detection", {})
            self.tolerance.set(str(detection.get("tolerance", self.tolerance.get())))
            self.interval_ms.set(str(detection.get("interval_ms", self.interval_ms.get())))
            self.cooldown_ms.set(str(detection.get("cooldown_ms", self.cooldown_ms.get())))
            self.silence_ms.set(str(detection.get("silence_ms", self.silence_ms.get())))

            tone = data.get("tone", {})
            self.beep_freq.set(str(tone.get("beep_freq", self.beep_freq.get())))
            self.beep_dur.set(str(tone.get("beep_dur", self.beep_dur.get())))

            if show_message:
                messagebox.showinfo("Profile loaded", f"Loaded from {self.config_path.name}")
        except Exception as exc:
            messagebox.showerror("Load failed", str(exc))

    def parse_zone(self) -> WatchZone:
        x = int(self.zone_x.get())
        y = int(self.zone_y.get())
        w = int(self.zone_w.get())
        h = int(self.zone_h.get())
        if w <= 0 or h <= 0:
            raise ValueError("Width and height must be > 0")
        return WatchZone(x=x, y=y, width=w, height=h)

    def parse_targets(self) -> list[tuple[int, int, int]]:
        targets: list[tuple[int, int, int]] = []
        for i in range(3):
            parsed = self._parse_color_slot(i)
            if parsed is not None:
                targets.append(parsed)
        if not targets:
            raise ValueError("Enter at least one target color")
        return targets

    def select_zone_overlay(self):
        self.root.withdraw()
        overlay = tk.Toplevel()
        overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.25)
        overlay.configure(bg="black")
        overlay.lift()
        overlay.attributes("-topmost", True)

        canvas = tk.Canvas(overlay, cursor="cross", highlightthickness=0)
        canvas.pack(fill="both", expand=True)

        start = {"x": 0, "y": 0}
        rect_id = {"value": None}

        def on_press(event):
            start["x"] = event.x
            start["y"] = event.y
            if rect_id["value"] is not None:
                canvas.delete(rect_id["value"])
            rect_id["value"] = canvas.create_rectangle(
                event.x,
                event.y,
                event.x,
                event.y,
                outline="red",
                width=2,
            )

        def on_drag(event):
            if rect_id["value"] is not None:
                canvas.coords(rect_id["value"], start["x"], start["y"], event.x, event.y)

        def on_release(event):
            x1, y1 = start["x"], start["y"]
            x2, y2 = event.x, event.y
            left, top = min(x1, x2), min(y1, y2)
            width, height = abs(x2 - x1), abs(y2 - y1)
            if width < 3 or height < 3:
                overlay.destroy()
                self.root.deiconify()
                return
            self.zone_x.set(str(left))
            self.zone_y.set(str(top))
            self.zone_w.set(str(width))
            self.zone_h.set(str(height))
            overlay.destroy()
            self.root.deiconify()

        def on_escape(_event):
            overlay.destroy()
            self.root.deiconify()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        overlay.bind("<Escape>", on_escape)

    def start(self):
        if self.running:
            return

        try:
            _ = self.parse_zone()
            _ = self.parse_targets()
            tolerance = int(self.tolerance.get())
            interval = int(self.interval_ms.get())
            freq = int(self.beep_freq.get())
            dur = int(self.beep_dur.get())
            cooldown = int(self.cooldown_ms.get())
            silence = int(self.silence_ms.get())
        except ValueError as exc:
            messagebox.showerror("Invalid settings", str(exc))
            return

        if not 0 <= tolerance <= 255:
            messagebox.showerror("Invalid settings", "Tolerance must be in 0..255")
            return
        if interval <= 0 or dur <= 0 or cooldown < 0:
            messagebox.showerror("Invalid settings", "Interval/duration must be > 0 and cooldown >= 0")
            return
        if silence <= 0:
            messagebox.showerror("Invalid settings", "Silence ms must be > 0")
            return
        if not 37 <= freq <= 32767:
            messagebox.showerror("Invalid settings", "Beep frequency must be in 37..32767 Hz")
            return

        self._init_capture_backend()
        if self.capture_backend == "none":
            messagebox.showerror(
                "Missing dependency",
                "No screen capture backend is available. Install 'mss' or install dxcam with numpy.",
            )
            return
        self.mute_until_ms = 0.0
        self.save_profile(show_message=False)

        self.running = True
        self.status_text.set(f"Status: Monitoring ({self.capture_backend})...")
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()

    def _init_capture_backend(self):
        self.capture_backend = "mss" if mss is not None else "none"
        self.dx_camera = None

        if dxcam is None or np is None:
            return

        try:
            self.dx_camera = dxcam.create(output_color="RGB")
            if self.dx_camera is not None:
                self.capture_backend = "dxcam"
        except Exception:
            self.dx_camera = None

    def stop(self):
        self.running = False
        self.status_text.set("Status: Stopped")

    def test_tone(self):
        try:
            freq = int(self.beep_freq.get())
            dur = int(self.beep_dur.get())
        except ValueError:
            messagebox.showerror("Invalid settings", "Beep frequency/duration must be numbers")
            return

        if not 37 <= freq <= 32767:
            messagebox.showerror("Invalid settings", "Beep frequency must be in 37..32767 Hz")
            return
        if dur <= 0:
            messagebox.showerror("Invalid settings", "Beep duration must be > 0")
            return

        tone_mode = self._play_tone(freq, dur)
        if tone_mode == "beep":
            self.status_text.set("Status: Test tone played.")
        else:
            self.status_text.set("Status: Test system alert played.")

    def _play_tone(self, freq: int, dur: int) -> str:
        try:
            winsound.Beep(freq, dur)
            return "beep"
        except RuntimeError:
            winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            return "message"

    def on_close(self):
        self._stop_preview()
        self.stop()
        self.save_profile(show_message=False)
        self.root.destroy()

    def silence_for_period(self):
        try:
            silence_ms = int(self.silence_ms.get())
        except ValueError:
            messagebox.showerror("Invalid settings", "Silence ms must be a number")
            return
        if silence_ms <= 0:
            messagebox.showerror("Invalid settings", "Silence ms must be > 0")
            return

        now = time.time() * 1000
        self.mute_until_ms = max(self.mute_until_ms, now + silence_ms)
        self.status_text.set(f"Status: Silenced for {silence_ms // 1000}s")

    def _capture_zone_image(self, zone: WatchZone) -> Image.Image:
        if self.capture_backend == "dxcam" and self.dx_camera is not None:
            right = zone.x + zone.width
            bottom = zone.y + zone.height
            frame = self.dx_camera.grab(region=(zone.x, zone.y, right, bottom))
            if frame is not None:
                return Image.fromarray(frame)

        if mss is None:
            raise RuntimeError("mss is not installed")

        monitor = {
            "left": zone.x,
            "top": zone.y,
            "width": zone.width,
            "height": zone.height,
        }
        with mss.mss() as sct:
            shot = sct.grab(monitor)
        return Image.frombytes("RGB", shot.size, shot.rgb)

    def open_preview_window(self):
        try:
            _ = self.parse_zone()
        except ValueError as exc:
            messagebox.showerror("Invalid zone", str(exc))
            return

        self._init_capture_backend()

        if self.preview_window is not None and self.preview_window.winfo_exists():
            self.preview_window.lift()
            return

        self.preview_window = tk.Toplevel(self.root)
        self.preview_window.title("Capture Preview")
        self.preview_window.geometry("520x420")

        self.preview_info = tk.StringVar(value=f"Backend: {self.capture_backend}")
        tk.Label(self.preview_window, textvariable=self.preview_info, anchor="w").pack(fill="x", padx=8, pady=(8, 4))

        self.preview_label = tk.Label(self.preview_window)
        self.preview_label.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.preview_window.protocol("WM_DELETE_WINDOW", self._stop_preview)
        self._update_preview_frame()

    def _update_preview_frame(self):
        if self.preview_window is None or not self.preview_window.winfo_exists() or self.preview_label is None:
            return

        try:
            zone = self.parse_zone()
            image = self._capture_zone_image(zone)
            image.thumbnail((1000, 700), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(image)
            self.preview_label.configure(image=photo)
            self.preview_label.image = photo
            if self.preview_info is not None:
                self.preview_info.set(f"Backend: {self.capture_backend} | Zone: {zone.width}x{zone.height}")
        except Exception as exc:
            if self.preview_info is not None:
                self.preview_info.set(f"Preview error: {exc}")

        self.preview_after_id = self.root.after(180, self._update_preview_frame)

    def _stop_preview(self):
        if self.preview_after_id is not None:
            self.root.after_cancel(self.preview_after_id)
            self.preview_after_id = None
        if self.preview_window is not None and self.preview_window.winfo_exists():
            self.preview_window.destroy()
        self.preview_window = None
        self.preview_label = None
        self.preview_info = None

    def _set_status(self, msg: str) -> None:
        self.root.after(0, self.status_text.set, msg)

    def monitor_loop(self):
        last_beep_ts = 0.0
        sct = None

        if self.capture_backend == "mss" and mss is not None:
            sct = mss.mss()

        try:
            zone = self.parse_zone()
            targets = self.parse_targets()
            tolerance = int(self.tolerance.get())
            interval_ms = int(self.interval_ms.get())
            freq = int(self.beep_freq.get())
            dur = int(self.beep_dur.get())
            cooldown_ms = int(self.cooldown_ms.get())
            silence_ms = int(self.silence_ms.get())

            while self.running:
                match_count = 0
                match_idx = -1

                if self.capture_backend == "dxcam" and self.dx_camera is not None:
                    right = zone.x + zone.width
                    bottom = zone.y + zone.height
                    frame = self.dx_camera.grab(region=(zone.x, zone.y, right, bottom))
                    match_count, match_idx = rgb_frame_best_match(frame, targets, tolerance)
                else:
                    monitor = {
                        "left": zone.x,
                        "top": zone.y,
                        "width": zone.width,
                        "height": zone.height,
                    }
                    shot = sct.grab(monitor)
                    match_count, match_idx = bgra_buffer_best_match(shot.raw, targets, tolerance)

                found = match_count > 0
                now = time.time() * 1000
                muted = now < self.mute_until_ms
                muted_seconds_left = max(0, int((self.mute_until_ms - now + 999) // 1000))
                color_i = match_idx + 1

                if found and not muted and now - last_beep_ts >= cooldown_ms:
                    tone_mode = self._play_tone(freq, dur)
                    last_beep_ts = now
                    suffix = "Tone played." if tone_mode == "beep" else "System alert played."
                    self._set_status(f"Status: Color {color_i} found ({match_count} px). {suffix}")
                elif found and muted:
                    self._set_status(f"Status: Color {color_i} found ({match_count} px, muted {muted_seconds_left}s left).")
                elif found:
                    self._set_status(f"Status: Color {color_i} found ({match_count} px), waiting cooldown...")
                else:
                    self._set_status(f"Status: Monitoring ({self.capture_backend})...")

                sleep_ms = min(interval_ms, silence_ms) if muted else interval_ms
                time.sleep(sleep_ms / 1000)

        except Exception as exc:
            self.running = False
            self._set_status(f"Status: Error - {exc}")
        finally:
            if sct is not None:
                sct.close()


def main():
    root = tk.Tk()
    app = App(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
