import ctypes
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
from PIL import Image, ImageDraw, ImageTk

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
        self.root.geometry("660x490")
        self.root.minsize(580, 460)
        self.root.resizable(True, True)
        self.config_path = self._resolve_profile_path()

        self.running = False
        self._restart_pending_id: str | None = None
        self.monitor_thread: threading.Thread | None = None
        self.capture_backend = "mss"
        self.dx_camera = None
        self.mute_until_ms = 0.0
        self.preview_window: tk.Toplevel | None = None
        self.preview_label: tk.Label | None = None
        self.preview_info: tk.StringVar | None = None
        self.preview_after_id: str | None = None
        self._preview_zone_num: int = 1

        self.zone1_enabled = tk.BooleanVar(value=True)
        self.zone_x = tk.StringVar(value="0")
        self.zone_y = tk.StringVar(value="0")
        self.zone_w = tk.StringVar(value="300")
        self.zone_h = tk.StringVar(value="200")

        self.zone2_enabled = tk.BooleanVar(value=False)
        self.zone2_x = tk.StringVar(value="0")
        self.zone2_y = tk.StringVar(value="0")
        self.zone2_w = tk.StringVar(value="300")
        self.zone2_h = tk.StringVar(value="200")

        self.zone3_enabled = tk.BooleanVar(value=False)
        self.zone3_x = tk.StringVar(value="0")
        self.zone3_y = tk.StringVar(value="0")
        self.zone3_w = tk.StringVar(value="300")
        self.zone3_h = tk.StringVar(value="200")

        self.color_hex_vars = [
            tk.StringVar(value="#7F3107"),
            tk.StringVar(value=""),
            tk.StringVar(value=""),
        ]

        self.tolerance = tk.StringVar(value="15")
        self.interval_ms = tk.StringVar(value="50")
        self.cooldown_ms = tk.StringVar(value="2000")
        self.silence_ms = tk.StringVar(value="60")
        self.dark_mode = tk.BooleanVar(value=True)

        self.status_text = tk.StringVar(value="Status: Idle")

        self._build_ui()
        self.load_profile(show_message=False)
        self._set_icon()
        self._apply_theme()
        self.dark_mode.trace_add("write", lambda *_: self._apply_theme())

        # Auto-restart traces — attached after load_profile so initial load doesn't trigger restart
        _watched = [
            self.zone1_enabled, self.zone_x, self.zone_y, self.zone_w, self.zone_h,
            self.zone2_enabled, self.zone2_x, self.zone2_y, self.zone2_w, self.zone2_h,
            self.zone3_enabled, self.zone3_x, self.zone3_y, self.zone3_w, self.zone3_h,
            self.color_hex_vars[0], self.color_hex_vars[1], self.color_hex_vars[2],
            self.tolerance, self.interval_ms, self.cooldown_ms, self.silence_ms,
        ]
        for _v in _watched:
            _v.trace_add("write", self._on_setting_changed)

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
        frame = tk.Frame(self.root, padx=8, pady=3)
        frame.pack(fill="both", expand=True)

        # --- 3 Watch Zones side by side ---
        zones_frame = tk.Frame(frame)
        zones_frame.pack(fill="x", pady=(0, 3))
        zones_frame.columnconfigure(0, weight=1)
        zones_frame.columnconfigure(1, weight=1)
        zones_frame.columnconfigure(2, weight=1)

        zone_configs = [
            ("Watch Zone 1", self.zone1_enabled, self.zone_x, self.zone_y, self.zone_w, self.zone_h,
             self.select_zone_overlay, lambda: self.open_preview_window(1)),
            ("Watch Zone 2", self.zone2_enabled, self.zone2_x, self.zone2_y, self.zone2_w, self.zone2_h,
             self.select_zone2_overlay, lambda: self.open_preview_window(2)),
            ("Watch Zone 3", self.zone3_enabled, self.zone3_x, self.zone3_y, self.zone3_w, self.zone3_h,
             self.select_zone3_overlay, lambda: self.open_preview_window(3)),
        ]

        for col, (title, enabled_var, xv, yv, wv, hv, sel_cmd, prev_cmd) in enumerate(zone_configs):
            box = tk.LabelFrame(zones_frame, text=title, padx=4, pady=2)
            box.grid(row=0, column=col, sticky="nsew", padx=(0, 0 if col == 2 else 4))
            box.columnconfigure(1, weight=1)
            box.columnconfigure(3, weight=1)
            tk.Checkbutton(box, text="Enabled", variable=enabled_var).grid(
                row=0, column=0, columnspan=4, sticky="w"
            )
            # X and Y on same row
            tk.Label(box, text="X").grid(row=1, column=0, sticky="w", padx=(0, 2))
            tk.Entry(box, textvariable=xv, width=6).grid(row=1, column=1, sticky="we", padx=(0, 4))
            tk.Label(box, text="Y").grid(row=1, column=2, sticky="w", padx=(0, 2))
            tk.Entry(box, textvariable=yv, width=6).grid(row=1, column=3, sticky="we")
            # W and H on same row
            tk.Label(box, text="W").grid(row=2, column=0, sticky="w", padx=(0, 2))
            tk.Entry(box, textvariable=wv, width=6).grid(row=2, column=1, sticky="we", padx=(0, 4))
            tk.Label(box, text="H").grid(row=2, column=2, sticky="w", padx=(0, 2))
            tk.Entry(box, textvariable=hv, width=6).grid(row=2, column=3, sticky="we")
            tk.Button(box, text="Select Zone", command=sel_cmd).grid(
                row=3, column=0, columnspan=4, sticky="we", pady=(3, 1)
            )
            tk.Button(box, text="Preview", command=prev_cmd).grid(
                row=4, column=0, columnspan=4, sticky="we", pady=(0, 0)
            )

        # --- Target Colors ---
        color_box = tk.LabelFrame(frame, text="Target Colors (up to 3)", padx=6, pady=3)
        color_box.pack(fill="x", pady=(0, 3))
        color_box.columnconfigure(1, weight=1)

        for i in range(3):
            r = i + 1
            tk.Label(color_box, text=f"Color {r}").grid(row=i, column=0, sticky="w", padx=(0, 8), pady=1)
            tk.Entry(color_box, textvariable=self.color_hex_vars[i]).grid(row=i, column=1, sticky="we", padx=2, pady=1)
            tk.Button(color_box, text="Pick", command=lambda idx=i: self.pick_color_from_screen(idx), width=5).grid(
                row=i, column=2, padx=(4, 0), pady=1
            )

        # --- Detection ---
        settings_box = tk.LabelFrame(frame, text="Detection", padx=6, pady=3)
        settings_box.pack(fill="x", pady=(0, 3))
        settings_box.columnconfigure(1, weight=1)
        settings_box.columnconfigure(3, weight=1)

        # Row 0: Tolerance | Scan interval
        tk.Label(settings_box, text="Tolerance (0-255)").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=1)
        tk.Entry(settings_box, textvariable=self.tolerance, width=7).grid(row=0, column=1, sticky="we", padx=(0, 12), pady=1)
        tk.Label(settings_box, text="Scan interval ms").grid(row=0, column=2, sticky="w", padx=(0, 4), pady=1)
        tk.Entry(settings_box, textvariable=self.interval_ms, width=7).grid(row=0, column=3, sticky="we", pady=1)

        # Row 1: Cooldown | Silence
        tk.Label(settings_box, text="Cooldown ms").grid(row=1, column=0, sticky="w", padx=(0, 4), pady=1)
        tk.Entry(settings_box, textvariable=self.cooldown_ms, width=7).grid(row=1, column=1, sticky="we", padx=(0, 12), pady=1)
        tk.Label(settings_box, text="Silence sec").grid(row=1, column=2, sticky="w", padx=(0, 4), pady=1)
        tk.Entry(settings_box, textvariable=self.silence_ms, width=7).grid(row=1, column=3, sticky="we", pady=1)

        # --- Controls ---
        controls = tk.Frame(frame)
        controls.pack(fill="x", pady=(0, 2))
        tk.Button(controls, text="Start", command=self.start).pack(side="left", fill="x", expand=True, padx=(0, 4))
        tk.Button(controls, text="Stop", command=self.stop).pack(side="left", fill="x", expand=True, padx=(0, 4))
        tk.Button(controls, text="Silence Now", command=self.silence_for_period).pack(
            side="left", fill="x", expand=True, padx=(0, 0)
        )

        profile_controls = tk.Frame(frame)
        profile_controls.pack(fill="x", pady=(0, 0))
        tk.Button(profile_controls, text="Save Profile", command=self.save_profile).pack(
            side="left", fill="x", expand=True, padx=(0, 4)
        )
        tk.Button(profile_controls, text="Load Profile", command=self.load_profile).pack(
            side="left", fill="x", expand=True, padx=(0, 4)
        )
        tk.Checkbutton(profile_controls, text="Dark Mode", variable=self.dark_mode).pack(
            side="left", padx=(0, 0)
        )

        tk.Label(frame, textvariable=self.status_text, anchor="w").pack(fill="x", pady=(3, 0))

    def _row(self, parent: tk.Widget, label: str, var: tk.StringVar, row: int):
        tk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=2)
        tk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="we", pady=2)
        parent.grid_columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    # Icon
    # ------------------------------------------------------------------
    def _set_icon(self):
        import math
        size = 64
        img = Image.new("RGBA", (size, size), (10, 12, 20, 255))
        draw = ImageDraw.Draw(img)
        cx, cy = size // 2, size // 2
        cyan = (0, 180, 216, 255)
        bright = (0, 230, 255, 255)

        # Outer ring
        r = 28
        draw.ellipse([cx - r, cy - r, cx + r, cy + r], outline=cyan, width=2)

        # Inner ring
        r2 = 9
        draw.ellipse([cx - r2, cy - r2, cx + r2, cy + r2], outline=(0, 180, 216, 160), width=1)

        # Crosshair lines with gap
        gap = 13
        draw.line([cx - r + 3, cy, cx - gap, cy], fill=cyan, width=1)
        draw.line([cx + gap, cy, cx + r - 3, cy], fill=cyan, width=1)
        draw.line([cx, cy - r + 3, cx, cy - gap], fill=cyan, width=1)
        draw.line([cx, cy + gap, cx, cy + r - 3], fill=cyan, width=1)

        # Cardinal tick marks on outer ring
        for deg in (0, 90, 180, 270):
            angle = math.radians(deg)
            px = cx + int(r * math.cos(angle))
            py = cy + int(r * math.sin(angle))
            draw.ellipse([px - 2, py - 2, px + 2, py + 2], fill=bright)

        # Centre dot
        draw.ellipse([cx - 2, cy - 2, cx + 2, cy + 2], fill=bright)

        self._icon_img = ImageTk.PhotoImage(img)
        self.root.iconphoto(True, self._icon_img)

    # ------------------------------------------------------------------
    # Dark / light theme
    # ------------------------------------------------------------------
    def _apply_theme(self):
        dark = self.dark_mode.get()
        if dark:
            s = {
                "root_bg":   "#0a0d12",
                "frame_bg":  "#0a0d12",
                "lframe_bg": "#0f1520",
                "lframe_fg": "#00b4d8",
                "label_fg":  "#8cb8cc",
                "entry_bg":  "#070a0f",
                "entry_fg":  "#c6dde8",
                "entry_sel": "#1e4060",
                "btn_bg":    "#162035",
                "btn_fg":    "#7ab8cc",
                "btn_act":   "#1e2f4a",
                "chk_sel":   "#09203f",
            }
        else:
            s = {
                "root_bg":   "SystemButtonFace",
                "frame_bg":  "SystemButtonFace",
                "lframe_bg": "SystemButtonFace",
                "lframe_fg": "SystemButtonText",
                "label_fg":  "SystemButtonText",
                "entry_bg":  "SystemWindow",
                "entry_fg":  "SystemWindowText",
                "entry_sel": "SystemHighlight",
                "btn_bg":    "SystemButtonFace",
                "btn_fg":    "SystemButtonText",
                "btn_act":   "SystemHighlight",
                "chk_sel":   "SystemWindow",
            }

        self.root.configure(bg=s["root_bg"])

        def apply(w, pbg):
            cls = w.__class__.__name__
            my_bg = pbg
            try:
                if cls == "LabelFrame":
                    my_bg = s["lframe_bg"]
                    w.configure(bg=my_bg, fg=s["lframe_fg"])
                elif cls == "Frame":
                    my_bg = pbg
                    w.configure(bg=my_bg)
                elif cls == "Label":
                    w.configure(bg=pbg, fg=s["label_fg"])
                elif cls == "Entry":
                    w.configure(
                        bg=s["entry_bg"], fg=s["entry_fg"],
                        insertbackground=s["entry_fg"],
                        selectbackground=s["entry_sel"],
                    )
                elif cls == "Button":
                    w.configure(
                        bg=s["btn_bg"], fg=s["btn_fg"],
                        activebackground=s["btn_act"],
                        activeforeground=s["btn_fg"],
                        relief="flat" if dark else "raised",
                        borderwidth=1,
                    )
                elif cls == "Checkbutton":
                    w.configure(
                        bg=pbg, fg=s["label_fg"],
                        activebackground=pbg,
                        selectcolor=s["chk_sel"],
                    )
            except tk.TclError:
                pass
            for child in w.winfo_children():
                apply(child, my_bg)

        for child in self.root.winfo_children():
            apply(child, s["root_bg"])

        # Dark title bar via Windows DWM API
        try:
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            # winfo_id() is the child Tk window; GetParent() gives the real frame HWND
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            if not hwnd:
                hwnd = self.root.winfo_id()
            val = ctypes.c_int(1 if dark else 0)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(val), ctypes.sizeof(val)
            )
            # Force a redraw so the titlebar updates immediately
            self.root.withdraw()
            self.root.deiconify()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Auto-restart on setting change
    # ------------------------------------------------------------------
    def _on_setting_changed(self, *_args):
        if not self.running:
            return
        if self._restart_pending_id is not None:
            self.root.after_cancel(self._restart_pending_id)
        self._restart_pending_id = self.root.after(700, self._do_restart)

    def _do_restart(self):
        self._restart_pending_id = None
        if not self.running:
            return
        self.stop()
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=0.5)
        self.start(silent=True)

    def _parse_color_slot(self, idx: int) -> tuple[int, int, int] | None:
        hex_value = self.color_hex_vars[idx].get().strip()
        if not hex_value:
            return None
        return hex_to_rgb(hex_value)

    def _get_profile_dict(self) -> dict:
        return {
            "zone": {
                "enabled": self.zone1_enabled.get(),
                "x": self.zone_x.get(),
                "y": self.zone_y.get(),
                "width": self.zone_w.get(),
                "height": self.zone_h.get(),
            },
            "zone2": {
                "enabled": self.zone2_enabled.get(),
                "x": self.zone2_x.get(),
                "y": self.zone2_y.get(),
                "width": self.zone2_w.get(),
                "height": self.zone2_h.get(),
            },
            "zone3": {
                "enabled": self.zone3_enabled.get(),
                "x": self.zone3_x.get(),
                "y": self.zone3_y.get(),
                "width": self.zone3_w.get(),
                "height": self.zone3_h.get(),
            },
            "color": {
                "colors": [{"hex": self.color_hex_vars[i].get()} for i in range(3)]
            },
            "detection": {
                "tolerance": self.tolerance.get(),
                "interval_ms": self.interval_ms.get(),
                "cooldown_ms": self.cooldown_ms.get(),
                "silence_sec": self.silence_ms.get(),
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
            self.zone1_enabled.set(bool(zone.get("enabled", self.zone1_enabled.get())))
            self.zone_x.set(str(zone.get("x", self.zone_x.get())))
            self.zone_y.set(str(zone.get("y", self.zone_y.get())))
            self.zone_w.set(str(zone.get("width", self.zone_w.get())))
            self.zone_h.set(str(zone.get("height", self.zone_h.get())))

            zone2 = data.get("zone2", {})
            self.zone2_enabled.set(bool(zone2.get("enabled", self.zone2_enabled.get())))
            self.zone2_x.set(str(zone2.get("x", self.zone2_x.get())))
            self.zone2_y.set(str(zone2.get("y", self.zone2_y.get())))
            self.zone2_w.set(str(zone2.get("width", self.zone2_w.get())))
            self.zone2_h.set(str(zone2.get("height", self.zone2_h.get())))

            zone3 = data.get("zone3", {})
            self.zone3_enabled.set(bool(zone3.get("enabled", self.zone3_enabled.get())))
            self.zone3_x.set(str(zone3.get("x", self.zone3_x.get())))
            self.zone3_y.set(str(zone3.get("y", self.zone3_y.get())))
            self.zone3_w.set(str(zone3.get("width", self.zone3_w.get())))
            self.zone3_h.set(str(zone3.get("height", self.zone3_h.get())))

            color = data.get("color", {})
            colors = color.get("colors")
            if isinstance(colors, list):
                for i in range(3):
                    item = colors[i] if i < len(colors) and isinstance(colors[i], dict) else {}
                    self.color_hex_vars[i].set(str(item.get("hex", self.color_hex_vars[i].get())))

            detection = data.get("detection", {})
            self.tolerance.set(str(detection.get("tolerance", self.tolerance.get())))
            self.interval_ms.set(str(detection.get("interval_ms", self.interval_ms.get())))
            self.cooldown_ms.set(str(detection.get("cooldown_ms", self.cooldown_ms.get())))
            self.silence_ms.set(str(detection.get("silence_sec", self.silence_ms.get())))

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

    def select_zone2_overlay(self):
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
                event.x, event.y, event.x, event.y, outline="cyan", width=2,
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
            self.zone2_x.set(str(left))
            self.zone2_y.set(str(top))
            self.zone2_w.set(str(width))
            self.zone2_h.set(str(height))
            overlay.destroy()
            self.root.deiconify()

        def on_escape(_event):
            overlay.destroy()
            self.root.deiconify()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        overlay.bind("<Escape>", on_escape)

    def pick_color_from_screen(self, idx: int):
        if mss is None:
            messagebox.showerror("Not available", "mss is required for color picking")
            return

        with mss.mss() as sct:
            monitor = sct.monitors[1]
            shot = sct.grab(monitor)
            screenshot = Image.frombytes("RGB", shot.size, shot.rgb)
            offset_x = monitor["left"]
            offset_y = monitor["top"]

        self.root.withdraw()
        overlay = tk.Toplevel()
        overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.01)
        overlay.configure(bg="black")
        overlay.attributes("-topmost", True)
        overlay.config(cursor="crosshair")

        def on_click(event):
            x = event.x_root - offset_x
            y = event.y_root - offset_y
            overlay.destroy()
            self.root.deiconify()
            try:
                r, g, b = screenshot.getpixel((x, y))
                self.color_hex_vars[idx].set(f"#{r:02X}{g:02X}{b:02X}")
            except Exception as exc:
                messagebox.showerror("Pick failed", str(exc))

        def on_escape(_event):
            overlay.destroy()
            self.root.deiconify()

        overlay.bind("<ButtonPress-1>", on_click)
        overlay.bind("<Escape>", on_escape)

    def _parse_zone2(self) -> WatchZone:
        x = int(self.zone2_x.get())
        y = int(self.zone2_y.get())
        w = int(self.zone2_w.get())
        h = int(self.zone2_h.get())
        if w <= 0 or h <= 0:
            raise ValueError("Zone 2 width and height must be > 0")
        return WatchZone(x=x, y=y, width=w, height=h)

    def _parse_zone3(self) -> WatchZone:
        x = int(self.zone3_x.get())
        y = int(self.zone3_y.get())
        w = int(self.zone3_w.get())
        h = int(self.zone3_h.get())
        if w <= 0 or h <= 0:
            raise ValueError("Zone 3 width and height must be > 0")
        return WatchZone(x=x, y=y, width=w, height=h)

    def select_zone3_overlay(self):
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
                event.x, event.y, event.x, event.y, outline="yellow", width=2,
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
            self.zone3_x.set(str(left))
            self.zone3_y.set(str(top))
            self.zone3_w.set(str(width))
            self.zone3_h.set(str(height))
            overlay.destroy()
            self.root.deiconify()

        def on_escape(_event):
            overlay.destroy()
            self.root.deiconify()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        overlay.bind("<Escape>", on_escape)

    def _match_zone(self, zone: WatchZone, targets: list, tolerance: int, sct) -> tuple[int, int]:
        if self.capture_backend == "dxcam" and self.dx_camera is not None:
            right = zone.x + zone.width
            bottom = zone.y + zone.height
            frame = self.dx_camera.grab(region=(zone.x, zone.y, right, bottom))
            return rgb_frame_best_match(frame, targets, tolerance)
        if sct is None:
            return 0, -1
        monitor = {"left": zone.x, "top": zone.y, "width": zone.width, "height": zone.height}
        shot = sct.grab(monitor)
        return bgra_buffer_best_match(shot.raw, targets, tolerance)

    def start(self, silent: bool = False):
        if self.running:
            return

        try:
            _ = self.parse_zone()
            _ = self.parse_targets()
            tolerance = int(self.tolerance.get())
            interval = int(self.interval_ms.get())

            cooldown = int(self.cooldown_ms.get())
            silence = int(self.silence_ms.get())
        except ValueError as exc:
            if not silent:
                messagebox.showerror("Invalid settings", str(exc))
            return

        if not 0 <= tolerance <= 255:
            if not silent:
                messagebox.showerror("Invalid settings", "Tolerance must be in 0..255")
            return
        if interval <= 0 or cooldown < 0:
            if not silent:
                messagebox.showerror("Invalid settings", "Interval must be > 0 and cooldown >= 0")
            return
        if silence <= 0:
            if not silent:
                messagebox.showerror("Invalid settings", "Silence ms must be > 0")
            return

        self._init_capture_backend()
        if self.capture_backend == "none":
            if not silent:
                messagebox.showerror(
                    "Missing dependency",
                    "No screen capture backend is available. Install 'mss' or install dxcam with numpy.",
                )
            return
        if not (self.zone1_enabled.get() or self.zone2_enabled.get() or self.zone3_enabled.get()):
            if not silent:
                messagebox.showerror("No zones", "Enable at least one Watch Zone before starting.")
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

    def _play_tone(self) -> str:
        try:
            winsound.Beep(1200, 180)
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
            silence_ms = int(self.silence_ms.get()) * 1000
        except ValueError:
            messagebox.showerror("Invalid settings", "Silence sec must be a number")
            return
        if silence_ms <= 0:
            messagebox.showerror("Invalid settings", "Silence sec must be > 0")
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

    def open_preview_window(self, zone_num: int = 1):
        self._preview_zone_num = zone_num
        try:
            if zone_num == 1:
                _ = self.parse_zone()
            elif zone_num == 2:
                _ = self._parse_zone2()
            else:
                _ = self._parse_zone3()
        except ValueError as exc:
            messagebox.showerror("Invalid zone", str(exc))
            return

        self._init_capture_backend()

        if self.preview_window is not None and self.preview_window.winfo_exists():
            self._preview_zone_num = zone_num
            self.preview_window.lift()
            return

        self.preview_window = tk.Toplevel(self.root)
        self.preview_window.title(f"Capture Preview (Zone {zone_num})")
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
            if self._preview_zone_num == 1:
                zone = self.parse_zone()
            elif self._preview_zone_num == 2:
                zone = self._parse_zone2()
            else:
                zone = self._parse_zone3()
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
        sct = mss.mss() if mss is not None else None

        try:
            active_zones: list[tuple[WatchZone, int]] = []
            if self.zone1_enabled.get():
                try:
                    active_zones.append((self.parse_zone(), 1))
                except ValueError:
                    pass
            if self.zone2_enabled.get():
                try:
                    active_zones.append((self._parse_zone2(), 2))
                except ValueError:
                    pass
            if self.zone3_enabled.get():
                try:
                    active_zones.append((self._parse_zone3(), 3))
                except ValueError:
                    pass
            if not active_zones:
                raise ValueError("No zones enabled — enable at least one Watch Zone before starting.")
            targets = self.parse_targets()
            tolerance = int(self.tolerance.get())
            interval_ms = int(self.interval_ms.get())
            cooldown_ms = int(self.cooldown_ms.get())
            silence_ms = int(self.silence_ms.get()) * 1000
            multi_zone = len(active_zones) > 1

            while self.running:
                match_count, match_idx, matched_zone = 0, -1, 1
                for zone, znum in active_zones:
                    c, idx = self._match_zone(zone, targets, tolerance, sct)
                    if c > match_count:
                        match_count, match_idx, matched_zone = c, idx, znum

                found = match_count > 0
                now = time.time() * 1000
                muted = now < self.mute_until_ms
                muted_seconds_left = max(0, int((self.mute_until_ms - now + 999) // 1000))
                color_i = match_idx + 1
                zone_label = f"Z{matched_zone} " if multi_zone else ""

                if found and not muted and now - last_beep_ts >= cooldown_ms:
                    tone_mode = self._play_tone()
                    last_beep_ts = now
                    suffix = "Tone played." if tone_mode == "beep" else "System alert played."
                    self._set_status(f"Status: {zone_label}Color {color_i} found ({match_count} px). {suffix}")
                elif found and muted:
                    self._set_status(f"Status: {zone_label}Color {color_i} found ({match_count} px, muted {muted_seconds_left}s left).")
                elif found:
                    self._set_status(f"Status: {zone_label}Color {color_i} found ({match_count} px), waiting cooldown...")
                else:
                    self._set_status(f"Status: Monitoring ({self.capture_backend})...")

                elapsed_ms = (time.time() * 1000) - now
                target_ms = min(interval_ms, silence_ms) if muted else interval_ms
                sleep_ms = max(0, target_ms - elapsed_ms)
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
