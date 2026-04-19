"""
SlapnMoan - Slap your laptop, it moans.
"""

import math
import random
import struct
import subprocess
import sys
import threading
import time
import tkinter as tk
import importlib.util
import shutil
import webbrowser
from pathlib import Path
from tkinter import ttk

try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    np = None
    NUMPY_AVAILABLE = False

try:
    import pygame
    PYGAME_AVAILABLE = True
except ImportError:
    pygame = None
    PYGAME_AVAILABLE = False

try:
    import sounddevice as sd
    MIC_AVAILABLE = True
except ImportError:
    MIC_AVAILABLE = False

WINRT_ACCEL = None
try:
    from winrt.windows.devices.sensors import Accelerometer as _A
    WINRT_ACCEL = _A
except ImportError:
    try:
        from winsdk.windows.devices.sensors import Accelerometer as _A
        WINRT_ACCEL = _A
    except ImportError:
        pass

try:
    from pynput import mouse as pynput_mouse
    PYNPUT_AVAILABLE = True
except ImportError:
    PYNPUT_AVAILABLE = False

import ctypes
from ctypes import byref
from ctypes.wintypes import DWORD

# APS / ShockMgr constants
GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
FILE_SHARE_READ = 0x00000001
FILE_SHARE_WRITE = 0x00000002
OPEN_EXISTING = 3
INVALID_HANDLE_VALUE = -1
IOCTL_APS_READ = 0x733FC

COOLDOWN_SEC = 1.2
SAMPLE_RATE = 44100
MIC_CHUNK = 1024
MIC_SAMPLERATE = 44100
MIC_THRESHOLD_DEF = 0.25
ACCEL_THRESHOLD = 2.8
MOUSE_SPEED_MIN = 3000

MOANS_DIR = Path(__file__).parent / "moans"

DETECTOR_MODES = [
    "Auto (all available)",
    "Mic only",
    "Accelerometer only",
    "Mouse jitter only",
    "Mic + accelerometer",
    "Mic + mouse jitter",
]

ACCEL_BACKENDS = ["Auto", "WinRT", "ThinkPad APS"]

DEPENDENCY_OPTIONS = [
    ("numpy", "numpy", "Core math"),
    ("pygame", "pygame", "Audio playback"),
    ("sounddevice", "sounddevice", "Mic capture"),
    ("pynput", "pynput", "Mouse jitter"),
    ("winrt-runtime", "winrt", "WinRT runtime"),
    ("winrt-Windows.Devices.Sensors", "winrt.windows.devices.sensors", "WinRT sensors"),
    ("winsdk", "winsdk.windows.devices.sensors", "WinSDK sensors"),
]
CORE_DEP_PACKAGES = {"numpy", "pygame", "sounddevice", "pynput"}
WINRT_DEP_PACKAGES = {"winrt-runtime", "winrt-Windows.Devices.Sensors"}
WINSDK_DEP_PACKAGES = {"winsdk"}


def load_moan_sounds() -> tuple[list, list]:
    sounds = []
    names = []
    if not PYGAME_AVAILABLE:
        return sounds, names

    if MOANS_DIR.exists():
        files = sorted(
            [
                f for f in MOANS_DIR.iterdir()
                if f.suffix.lower() in (".mp3", ".wav", ".ogg")
            ],
            key=lambda f: int(f.stem) if f.stem.isdigit() else 999,
        )

        for f in files:
            try:
                sounds.append(pygame.mixer.Sound(str(f)))
                names.append(f.name)
                print(f"[Sound] Loaded {f.name}")
            except Exception as e:
                print(f"[Sound] Failed to load {f.name}: {e}")

    if not sounds and NUMPY_AVAILABLE:
        print("[Sound] No valid moans found, using fallback synth tone")
        sounds = [_synth_fallback()]
        names = ["Synth Tone"]

    return sounds, names


def _synth_fallback():
    sr = 44100
    t = np.linspace(0, 0.9, int(sr * 0.9), dtype=np.float32)
    f = 300 * np.exp(-2.5 * t) + 12 * np.sin(2 * np.pi * 5.5 * t)
    p = np.cumsum(2 * np.pi * f / sr)
    s = np.sin(p) + 0.3 * np.sin(2 * p)
    n = int(0.03 * sr)
    e = np.concatenate([
        np.linspace(0, 1, n),
        np.exp(-4.0 * np.linspace(0, 1, len(t) - n)),
    ])
    out = (s * e * 0.72 * 32767).astype(np.int16)
    return pygame.sndarray.make_sound(np.column_stack([out, out]))


class SlapMoanApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.main_hidden_for_setup = False
        self.running = False
        self.last_moan = 0.0
        self.moan_count = 0

        self.accel_obj = None
        self.accel_token = None
        self.aps_handle = None
        self._aps_poller_thread = None

        self._mouse_listener = None
        self._prev_mouse_pos = None
        self._prev_mouse_t = 0.0

        self._audio_stream = None

        self.has_mic = False
        self.has_mouse = False
        self.has_winrt = False
        self.has_aps = False

        self._last_accel_backend = "None"
        self._last_mic_device = "None"
        self._last_mic_error = ""
        self.dep_vars = {}
        self.dep_state_labels = {}
        self.dep_checkbuttons = {}
        self.dep_checkbuttons = {}
        self.dep_status_var = tk.StringVar(value="Dependency installer: ready")
        self.installing_dependencies = False
        self.dep_dialog = None
        self.dep_install_btn = None
        self.dep_skip_btn = None
        self.dep_install_python_btn = None

        self.audio_ready = False
        if PYGAME_AVAILABLE:
            try:
                pygame.mixer.init(frequency=SAMPLE_RATE, size=-16, channels=2, buffer=512)
                self.audio_ready = True
            except Exception as e:
                print(f"[Audio] pygame mixer init failed: {e}")
        self.moan_sounds, self.sound_names = load_moan_sounds()

        self._build_ui()
        self._detect_and_badge()
        self.root.after(200, self._maybe_open_dependency_modal)

    def _build_ui(self):
        self.root.title("SlapnMoan")
        self.root.geometry("470x760")
        self.root.minsize(420, 620)
        self.root.resizable(True, True)
        self.root.configure(bg="#0d0d0d")

        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "SlapMeter.Horizontal.TProgressbar",
            troughcolor="#1a1a1a",
            background="#ff4d6d",
            darkcolor="#ff4d6d",
            lightcolor="#ff4d6d",
            bordercolor="#0d0d0d",
            relief="flat",
        )
        style.configure(
            "Slap.TCombobox",
            fieldbackground="#151515",
            background="#151515",
            foreground="#f2f2f2",
            bordercolor="#3b3b3b",
            arrowcolor="#ff4d6d",
            selectbackground="#2a2a2a",
            selectforeground="#ffffff",
        )
        style.map(
            "Slap.TCombobox",
            fieldbackground=[("readonly", "#151515")],
            foreground=[("readonly", "#f2f2f2")],
            selectbackground=[("readonly", "#2a2a2a")],
            selectforeground=[("readonly", "#ffffff")],
        )

        # Force readable dropdown list colors on Windows/Tk.
        self.root.option_add("*TCombobox*Listbox.background", "#151515")
        self.root.option_add("*TCombobox*Listbox.foreground", "#f2f2f2")
        self.root.option_add("*TCombobox*Listbox.selectBackground", "#2a2a2a")
        self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

        outer = tk.Frame(self.root, bg="#0d0d0d")
        outer.pack(fill="both", expand=True)

        self.scroll_canvas = tk.Canvas(
            outer,
            bg="#0d0d0d",
            highlightthickness=0,
            bd=0,
        )
        scroll = tk.Scrollbar(outer, orient="vertical", command=self.scroll_canvas.yview)
        self.scroll_canvas.configure(yscrollcommand=scroll.set)
        self.scroll_canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        panel = tk.Frame(self.scroll_canvas, bg="#0d0d0d")
        panel_window = self.scroll_canvas.create_window((0, 0), window=panel, anchor="nw")
        panel.bind(
            "<Configure>",
            lambda _e: self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all")),
        )
        self.scroll_canvas.bind(
            "<Configure>",
            lambda e: self.scroll_canvas.itemconfigure(panel_window, width=e.width),
        )
        self.scroll_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        tk.Label(
            panel,
            text="SlapnMoan",
            font=("Segoe UI", 24, "bold"),
            bg="#0d0d0d",
            fg="#ff4d6d",
        ).pack(pady=(20, 4))
        tk.Label(
            panel,
            text="Slap your laptop. It moans.",
            font=("Segoe UI", 11),
            bg="#0d0d0d",
            fg="#8a8a8a",
        ).pack()

        tk.Label(
            panel,
            text="DETECTED SENSORS",
            font=("Segoe UI", 8),
            bg="#0d0d0d",
            fg="#474747",
        ).pack(pady=(14, 6))

        badge_row = tk.Frame(panel, bg="#0d0d0d")
        badge_row.pack()
        self.badge_mic = self._badge(badge_row, "Mic")
        self.badge_accel = self._badge(badge_row, "Accelerometer")
        self.badge_mouse = self._badge(badge_row, "Mouse jitter")

        self.accel_info_var = tk.StringVar(value="Accel backend: not checked")
        tk.Label(
            panel,
            textvariable=self.accel_info_var,
            font=("Segoe UI", 8),
            bg="#0d0d0d",
            fg="#5f5f5f",
            wraplength=430,
            justify="center",
        ).pack(pady=(6, 2))

        controls = tk.Frame(panel, bg="#0d0d0d")
        controls.pack(pady=(8, 8), fill="x")

        tk.Label(
            controls,
            text="DETECTOR MODE",
            font=("Segoe UI", 8),
            bg="#0d0d0d",
            fg="#555",
        ).pack()
        self.detector_mode_var = tk.StringVar(value=DETECTOR_MODES[0])
        self.detector_mode_combo = ttk.Combobox(
            controls,
            textvariable=self.detector_mode_var,
            state="readonly",
            values=DETECTOR_MODES,
            width=34,
            style="Slap.TCombobox",
            font=("Segoe UI", 10),
        )
        self.detector_mode_combo.pack(pady=(2, 6))
        self.detector_mode_combo.bind("<<ComboboxSelected>>", self._on_detector_mode_change)

        tk.Label(
            controls,
            text="ACCEL BACKEND",
            font=("Segoe UI", 8),
            bg="#0d0d0d",
            fg="#555",
        ).pack()
        self.accel_backend_var = tk.StringVar(value=ACCEL_BACKENDS[0])
        self.accel_backend_combo = ttk.Combobox(
            controls,
            textvariable=self.accel_backend_var,
            state="readonly",
            values=ACCEL_BACKENDS,
            width=34,
            style="Slap.TCombobox",
            font=("Segoe UI", 10),
        )
        self.accel_backend_combo.pack(pady=(2, 2))
        self.accel_backend_combo.bind("<<ComboboxSelected>>", self._on_accel_backend_change)

        tk.Button(
            controls,
            text="Refresh sensor detection",
            font=("Segoe UI", 9),
            bg="#101010",
            fg="#bdbdbd",
            relief="flat",
            padx=12,
            pady=4,
            cursor="hand2",
            command=self._detect_and_badge,
        ).pack(pady=(8, 2))

        tk.Label(
            panel,
            text="IMPACT LEVEL",
            font=("Segoe UI", 9),
            bg="#0d0d0d",
            fg="#555",
        ).pack(pady=(12, 4))
        self.impact_var = tk.DoubleVar(value=0.0)
        ttk.Progressbar(
            panel,
            variable=self.impact_var,
            maximum=100,
            length=360,
            style="SlapMeter.Horizontal.TProgressbar",
        ).pack()
        self.impact_lbl = tk.Label(
            panel,
            text="-",
            font=("Consolas", 12, "bold"),
            bg="#0d0d0d",
            fg="#ff4d6d",
        )
        self.impact_lbl.pack(pady=4)

        slider_frame = tk.Frame(panel, bg="#0d0d0d")
        slider_frame.pack(pady=4)

        tk.Label(
            slider_frame,
            text="MIC SENSITIVITY (lower = more sensitive)",
            font=("Segoe UI", 8),
            bg="#0d0d0d",
            fg="#555",
        ).grid(row=0, column=0, sticky="w", padx=14)

        self.mic_thresh = tk.DoubleVar(value=MIC_THRESHOLD_DEF)
        tk.Scale(
            slider_frame,
            variable=self.mic_thresh,
            from_=0.02,
            to=0.8,
            resolution=0.01,
            orient=tk.HORIZONTAL,
            length=360,
            bg="#0d0d0d",
            fg="#ccc",
            troughcolor="#1a1a1a",
            highlightthickness=0,
            activebackground="#ff4d6d",
            font=("Segoe UI", 8),
        ).grid(row=1, column=0, padx=14)

        tk.Label(
            slider_frame,
            text="ACCEL THRESHOLD (g-force or APS delta)",
            font=("Segoe UI", 8),
            bg="#0d0d0d",
            fg="#555",
        ).grid(row=2, column=0, sticky="w", padx=14, pady=(6, 0))

        self.accel_thresh = tk.DoubleVar(value=ACCEL_THRESHOLD)
        tk.Scale(
            slider_frame,
            variable=self.accel_thresh,
            from_=1.5,
            to=6.0,
            resolution=0.1,
            orient=tk.HORIZONTAL,
            length=360,
            bg="#0d0d0d",
            fg="#ccc",
            troughcolor="#1a1a1a",
            highlightthickness=0,
            activebackground="#ff4d6d",
            font=("Segoe UI", 8),
        ).grid(row=3, column=0, padx=14)

        sounds_block = tk.Frame(panel, bg="#0d0d0d")
        sounds_block.pack(fill="x", padx=16, pady=(6, 4))

        self.sounds_var = tk.StringVar()
        tk.Label(
            sounds_block,
            textvariable=self.sounds_var,
            font=("Segoe UI", 9),
            bg="#0d0d0d",
            fg="#5b6678",
        ).pack(anchor="center")

        tk.Label(
            sounds_block,
            text="MOAN SOUND",
            font=("Segoe UI", 8),
            bg="#0d0d0d",
            fg="#555",
        ).pack(pady=(8, 2))

        self.current_moan_var = tk.StringVar()
        self.current_moan_combo = ttk.Combobox(
            sounds_block,
            textvariable=self.current_moan_var,
            state="readonly",
            width=36,
            style="Slap.TCombobox",
            font=("Segoe UI", 10),
        )
        self.current_moan_combo.pack()

        sound_actions = tk.Frame(sounds_block, bg="#0d0d0d")
        sound_actions.pack(pady=(8, 0))
        tk.Button(
            sound_actions,
            text="Preview selected",
            font=("Segoe UI", 9),
            bg="#151515",
            fg="#cfcfcf",
            relief="flat",
            cursor="hand2",
            command=self._test_moan,
            padx=10,
            pady=4,
        ).pack(side=tk.LEFT, padx=4)
        tk.Button(
            sound_actions,
            text="Reload sounds",
            font=("Segoe UI", 9),
            bg="#151515",
            fg="#cfcfcf",
            relief="flat",
            cursor="hand2",
            command=self._reload_sounds,
            padx=10,
            pady=4,
        ).pack(side=tk.LEFT, padx=4)

        self.status_var = tk.StringVar(value="Idle")
        self.status_lbl = tk.Label(
            panel,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bg="#0d0d0d",
            fg="#666",
            justify="center",
            wraplength=430,
        )
        self.status_lbl.pack(pady=(10, 2))

        self.count_var = tk.StringVar(value="Moans: 0")
        tk.Label(
            panel,
            textvariable=self.count_var,
            font=("Segoe UI", 9),
            bg="#0d0d0d",
            fg="#444",
        ).pack()

        button_frame = tk.Frame(panel, bg="#0d0d0d")
        button_frame.pack(pady=(14, 18))

        self.start_btn = tk.Button(
            button_frame,
            text="START",
            font=("Segoe UI", 11, "bold"),
            bg="#ff4d6d",
            fg="white",
            relief="flat",
            padx=20,
            pady=8,
            cursor="hand2",
            command=self.toggle,
        )
        self.start_btn.pack(side=tk.LEFT, padx=6)

        tk.Button(
            button_frame,
            text="TEST MOAN",
            font=("Segoe UI", 11),
            bg="#1a1a1a",
            fg="#ccc",
            relief="flat",
            padx=14,
            pady=8,
            cursor="hand2",
            command=self._test_moan,
        ).pack(side=tk.LEFT, padx=6)

        self._update_moan_combo()

    def _badge(self, parent, label):
        lbl = tk.Label(
            parent,
            text=f"o  {label}",
            font=("Segoe UI", 9),
            bg="#1a1a1a",
            fg="#555",
            padx=10,
            pady=5,
            relief="flat",
        )
        lbl.pack(side=tk.LEFT, padx=4)
        return lbl

    def _set_badge(self, badge, active: bool):
        txt = badge.cget("text")[3:]
        if active:
            badge.config(text=f"v  {txt}", fg="#4dff91", bg="#0a2018")
        else:
            badge.config(text=f"x  {txt}", fg="#804040", bg="#1a0d0d")

    def _on_mousewheel(self, event):
        if hasattr(self, "scroll_canvas") and self.scroll_canvas.winfo_exists():
            self.scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _module_installed(self, module_name: str) -> bool:
        try:
            return importlib.util.find_spec(module_name) is not None
        except Exception:
            return False

    def _python_available(self) -> bool:
        return Path(sys.executable).exists() or bool(shutil.which("py") or shutil.which("python"))

    def _dependency_snapshot(self) -> tuple[dict[str, bool], list[str], list[str], list[str], bool, bool]:
        installed = {}
        for pkg, module_name, _desc in DEPENDENCY_OPTIONS:
            installed[pkg] = self._module_installed(module_name)

        missing_core = [pkg for pkg in CORE_DEP_PACKAGES if not installed.get(pkg, False)]
        missing_any = [pkg for pkg, _module, _desc in DEPENDENCY_OPTIONS if not installed.get(pkg, False)]
        accel_ready = (
            self._module_installed("winrt.windows.devices.sensors")
            or self._module_installed("winsdk.windows.devices.sensors")
        )
        missing_required = list(missing_core)
        if not accel_ready:
            # Require at least one accel backend, not both stacks.
            missing_required.append("winsdk")
        python_ok = self._python_available()
        return installed, missing_core, missing_required, missing_any, accel_ready, python_ok

    def _maybe_open_dependency_modal(self):
        if self.dep_dialog is not None and self.dep_dialog.winfo_exists():
            return

        _installed, missing_core, missing_required, missing_any, accel_ready, python_ok = self._dependency_snapshot()
        if not missing_required:
            try:
                self.root.attributes("-alpha", 1.0)
            except Exception:
                pass
            self.main_hidden_for_setup = False
            self.root.deiconify()
            self.root.lift()
            return
        self._open_dependency_modal(missing_core, missing_required, missing_any, accel_ready, python_ok)

    def _open_dependency_modal(self, missing_core: list[str], missing_required: list[str], missing_any: list[str], accel_ready: bool, python_ok: bool):
        self.dep_dialog = tk.Toplevel(self.root)
        self.dep_dialog.title("Dependency Setup")
        self.dep_dialog.geometry("560x440")
        self.dep_dialog.resizable(False, False)
        self.dep_dialog.configure(bg="#111111")
        self.dep_dialog.transient(self.root)
        self.dep_dialog.grab_set()
        self.dep_dialog.protocol("WM_DELETE_WINDOW", self._skip_dependency_setup)
        self.dep_dialog.lift()
        self.dep_dialog.focus_force()
        try:
            self.root.attributes("-alpha", 0.0)
            self.main_hidden_for_setup = True
        except Exception:
            self.main_hidden_for_setup = False

        self.dep_vars = {}
        self.dep_state_labels = {}

        body = tk.Frame(self.dep_dialog, bg="#111111")
        body.pack(fill="both", expand=True, padx=16, pady=14)

        tk.Label(
            body,
            text="Set Up Dependencies",
            font=("Segoe UI", 14, "bold"),
            bg="#111111",
            fg="#ff4d6d",
        ).pack(anchor="w")

        summary = []
        if missing_required:
            summary.append(f"{len(missing_required)} required package(s) are missing.")
        if missing_core:
            summary.append("Core packages are missing.")
        if not accel_ready:
            summary.append("No accelerometer backend is installed yet.")
        if not python_ok:
            summary.append("Python runtime was not detected on this system.")
        summary_text = " ".join(summary) if summary else "Dependencies look good."

        tk.Label(
            body,
            text=summary_text + " Install now or skip and continue.",
            font=("Segoe UI", 9),
            bg="#111111",
            fg="#b0b0b0",
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(4, 10))

        dep_grid = tk.Frame(body, bg="#111111")
        dep_grid.pack(fill="x")

        selected_default = set(CORE_DEP_PACKAGES)
        if not accel_ready:
            selected_default |= WINSDK_DEP_PACKAGES

        for row, (pkg, _mod, desc) in enumerate(DEPENDENCY_OPTIONS):
            var = tk.BooleanVar(value=pkg in selected_default)
            self.dep_vars[pkg] = var

            cb = tk.Checkbutton(
                dep_grid,
                text=pkg,
                variable=var,
                bg="#111111",
                fg="#e0e0e0",
                selectcolor="#1c1c1c",
                activebackground="#111111",
                activeforeground="#ffffff",
                highlightthickness=0,
                font=("Segoe UI", 9),
                anchor="w",
            )
            cb.grid(row=row, column=0, sticky="w", pady=2)
            self.dep_checkbuttons[pkg] = cb

            tk.Label(
                dep_grid,
                text=desc,
                bg="#111111",
                fg="#818181",
                font=("Segoe UI", 8),
                anchor="w",
            ).grid(row=row, column=1, sticky="w", padx=(8, 8))

            lbl = tk.Label(
                dep_grid,
                text="checking...",
                bg="#111111",
                fg="#9a9a9a",
                font=("Segoe UI", 8),
                anchor="e",
            )
            lbl.grid(row=row, column=2, sticky="e")
            self.dep_state_labels[pkg] = lbl

        dep_grid.grid_columnconfigure(1, weight=1)

        preset_row = tk.Frame(body, bg="#111111")
        preset_row.pack(anchor="w", pady=(10, 0))
        tk.Button(
            preset_row,
            text="Preset: WinRT",
            font=("Segoe UI", 9),
            bg="#1a1a1a",
            fg="#e0e0e0",
            relief="flat",
            cursor="hand2",
            command=lambda: self._set_dependency_preset("winrt"),
            padx=10,
            pady=4,
        ).pack(side=tk.LEFT)
        tk.Button(
            preset_row,
            text="Preset: WinSDK",
            font=("Segoe UI", 9),
            bg="#1a1a1a",
            fg="#e0e0e0",
            relief="flat",
            cursor="hand2",
            command=lambda: self._set_dependency_preset("winsdk"),
            padx=10,
            pady=4,
        ).pack(side=tk.LEFT, padx=6)
        tk.Button(
            preset_row,
            text="Refresh status",
            font=("Segoe UI", 9),
            bg="#1a1a1a",
            fg="#e0e0e0",
            relief="flat",
            cursor="hand2",
            command=self._refresh_dependency_status,
            padx=10,
            pady=4,
        ).pack(side=tk.LEFT, padx=6)
        self.dep_install_python_btn = tk.Button(
            preset_row,
            text="Install Python (winget)",
            font=("Segoe UI", 9),
            bg="#1a1a1a",
            fg="#e0e0e0",
            relief="flat",
            cursor="hand2",
            command=self._install_python_winget,
            padx=10,
            pady=4,
        )
        self.dep_install_python_btn.pack(side=tk.LEFT, padx=6)
        tk.Button(
            preset_row,
            text="Open Python download",
            font=("Segoe UI", 9),
            bg="#1a1a1a",
            fg="#e0e0e0",
            relief="flat",
            cursor="hand2",
            command=self._open_python_download_page,
            padx=10,
            pady=4,
        ).pack(side=tk.LEFT)

        self.dep_status_var.set("Choose dependencies then click install, or skip for now.")
        tk.Label(
            body,
            textvariable=self.dep_status_var,
            font=("Segoe UI", 9),
            bg="#111111",
            fg="#95a6bc",
            wraplength=520,
            justify="left",
        ).pack(anchor="w", pady=(12, 8))

        action_row = tk.Frame(body, bg="#111111")
        action_row.pack(anchor="e")
        self.dep_skip_btn = tk.Button(
            action_row,
            text="Skip for now",
            font=("Segoe UI", 9),
            bg="#161616",
            fg="#d0d0d0",
            relief="flat",
            cursor="hand2",
            command=self._skip_dependency_setup,
            padx=12,
            pady=6,
        )
        self.dep_skip_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.dep_install_btn = tk.Button(
            action_row,
            text="Install selected",
            font=("Segoe UI", 9, "bold"),
            bg="#ff4d6d",
            fg="white",
            relief="flat",
            cursor="hand2",
            command=self._install_selected_dependencies,
            padx=12,
            pady=6,
        )
        self.dep_install_btn.pack(side=tk.LEFT)

        self._refresh_dependency_status()

    def _refresh_dependency_status(self):
        installed, missing_core, missing_required, missing_any, accel_ready, python_ok = self._dependency_snapshot()

        for pkg, _module_name, _desc in DEPENDENCY_OPTIONS:
            lbl = self.dep_state_labels.get(pkg)
            if lbl is None:
                continue
            ok = installed.get(pkg, False)
            lbl.config(text="installed" if ok else "missing", fg="#4dff91" if ok else "#ffb347")
            cb = self.dep_checkbuttons.get(pkg)
            if cb is not None:
                if ok:
                    self.dep_vars[pkg].set(False)
                    cb.config(state=tk.DISABLED, disabledforeground="#6f6f6f")
                else:
                    cb.config(state=tk.NORMAL)

        if not missing_required:
            self.dep_status_var.set("All dependencies required for full features are installed.")
        else:
            parts = []
            if missing_required:
                parts.append("Required missing: " + ", ".join(sorted(set(missing_required))))
            if missing_core:
                parts.append("Missing core: " + ", ".join(sorted(missing_core)))
            if not accel_ready:
                parts.append("No accel backend: install WinRT or WinSDK")
            if not python_ok:
                parts.append("Python runtime missing: use Install Python button")
            self.dep_status_var.set(" | ".join(parts))

    def _selected_packages(self) -> list[str]:
        installed, _missing_core, _missing_required, _missing_any, _accel_ready, _python_ok = self._dependency_snapshot()
        return [pkg for pkg, var in self.dep_vars.items() if var.get() and not installed.get(pkg, False)]

    def _set_dependency_preset(self, preset: str):
        selected = set(CORE_DEP_PACKAGES)
        if preset == "winrt":
            selected |= WINRT_DEP_PACKAGES
        elif preset == "winsdk":
            selected |= WINSDK_DEP_PACKAGES

        installed, _missing_core, _missing_required, _missing_any, _accel_ready, _python_ok = self._dependency_snapshot()
        for pkg, var in self.dep_vars.items():
            var.set((pkg in selected) and (not installed.get(pkg, False)))

        pretty = "WinRT" if preset == "winrt" else "WinSDK"
        self.dep_status_var.set(f"Selected {pretty} preset. Click install to apply.")

    def _python_for_pip(self) -> str:
        local_venv_python = Path(__file__).parent / ".venv" / "Scripts" / "python.exe"
        if local_venv_python.exists():
            return str(local_venv_python)
        return sys.executable

    def _install_selected_dependencies(self):
        if self.installing_dependencies:
            return

        packages = self._selected_packages()
        if not packages:
            self.dep_status_var.set("No packages selected.")
            return

        both_backends = ("winsdk" in packages) and (
            "winrt-runtime" in packages or "winrt-Windows.Devices.Sensors" in packages
        )
        if both_backends:
            self.dep_status_var.set(
                "Both WinRT and WinSDK selected. Install continues, but usually one backend is enough."
            )

        py = self._python_for_pip()
        self.installing_dependencies = True
        if self.dep_install_btn is not None:
            self.dep_install_btn.config(state=tk.DISABLED)
        if self.dep_skip_btn is not None:
            self.dep_skip_btn.config(state=tk.DISABLED)
        self.dep_status_var.set(
            f"Installing {len(packages)} package(s) using {py} ... this can take a minute."
        )
        threading.Thread(
            target=self._dependency_install_worker,
            args=(packages, py),
            daemon=True,
        ).start()

    def _dependency_install_worker(self, packages: list[str], py: str):
        cmd = [py, "-m", "pip", "install"] + packages
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False,
            )
            success = result.returncode == 0
            err_line = self._extract_install_error(result.stdout or "", result.stderr or "")
        except Exception as e:
            success = False
            err_line = str(e)

        try:
            self.root.after(0, self._finish_dependency_install, success, packages, err_line)
        except Exception:
            pass

    def _extract_install_error(self, stdout_text: str, stderr_text: str) -> str:
        lines = []
        lines.extend((stderr_text or "").splitlines())
        lines.extend((stdout_text or "").splitlines())
        ignored = (
            "this error originates from a subprocess",
            "to update, run:",
            "[notice]",
            "warning:",
        )

        for line in reversed(lines):
            t = line.strip()
            if not t:
                continue
            low = t.lower()
            if any(tok in low for tok in ignored):
                continue
            if "error" in low or "failed" in low or "could not" in low:
                return t

        for line in reversed(lines):
            t = line.strip()
            if t:
                return t
        return "pip install failed."

    def _open_python_download_page(self):
        try:
            webbrowser.open("https://www.python.org/downloads/windows/")
            self.dep_status_var.set("Opened Python downloads page in your browser.")
        except Exception as e:
            self.dep_status_var.set(f"Could not open browser: {e}")

    def _install_python_winget(self):
        if self.installing_dependencies:
            return
        if not shutil.which("winget"):
            self.dep_status_var.set("winget is not available. Use 'Open Python download' instead.")
            return

        self.installing_dependencies = True
        if self.dep_install_btn is not None:
            self.dep_install_btn.config(state=tk.DISABLED)
        if self.dep_skip_btn is not None:
            self.dep_skip_btn.config(state=tk.DISABLED)
        if self.dep_install_python_btn is not None:
            self.dep_install_python_btn.config(state=tk.DISABLED)
        self.dep_status_var.set("Installing Python via winget...")
        threading.Thread(target=self._python_install_worker, daemon=True).start()

    def _python_install_worker(self):
        cmd = [
            "winget",
            "install",
            "-e",
            "--id",
            "Python.Python.3.13",
            "--accept-package-agreements",
            "--accept-source-agreements",
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            success = result.returncode == 0
            err = self._extract_install_error(result.stdout or "", result.stderr or "")
        except Exception as e:
            success = False
            err = str(e)
        try:
            self.root.after(0, self._finish_python_install, success, err)
        except Exception:
            pass

    def _finish_python_install(self, success: bool, err_line: str):
        self.installing_dependencies = False
        if self.dep_install_btn is not None:
            self.dep_install_btn.config(state=tk.NORMAL)
        if self.dep_skip_btn is not None:
            self.dep_skip_btn.config(state=tk.NORMAL)
        if self.dep_install_python_btn is not None:
            self.dep_install_python_btn.config(state=tk.NORMAL)

        if success:
            self.dep_status_var.set("Python installed successfully. You may need to restart this app.")
        else:
            self.dep_status_var.set(
                "Python install failed: "
                + (err_line or "Unknown error")
                + " | Try 'Open Python download'."
            )
        self._refresh_dependency_status()

    def _finish_dependency_install(self, success: bool, packages: list[str], err_line: str):
        self.installing_dependencies = False
        if self.dep_install_btn is not None:
            self.dep_install_btn.config(state=tk.NORMAL)
        if self.dep_skip_btn is not None:
            self.dep_skip_btn.config(state=tk.NORMAL)
        if self.dep_install_python_btn is not None:
            self.dep_install_python_btn.config(state=tk.NORMAL)
        self._refresh_dependency_status()

        if success:
            installed, missing_core, missing_required, missing_any, accel_ready, python_ok = self._dependency_snapshot()
            if not missing_required:
                self.dep_status_var.set("Install complete. Dependencies look good, continuing to app...")
                self._detect_and_badge()
                self.root.after(400, self._close_dependency_modal)
            else:
                self.dep_status_var.set(
                    "Install complete for: "
                    + ", ".join(packages)
                    + ". Some items are still missing; you can install more or skip."
                )
        else:
            msg = err_line if err_line else "pip install failed."
            self.dep_status_var.set(f"Install failed: {msg}")

    def _skip_dependency_setup(self):
        self.dep_status_var.set("Dependency setup skipped. You can still use the app.")
        self._close_dependency_modal()

    def _close_dependency_modal(self):
        if self.dep_dialog is not None and self.dep_dialog.winfo_exists():
            try:
                self.dep_dialog.grab_release()
            except Exception:
                pass
            self.dep_dialog.destroy()
        self.dep_dialog = None
        self.dep_install_btn = None
        self.dep_skip_btn = None
        self.dep_install_python_btn = None
        try:
            if self.main_hidden_for_setup:
                self.root.attributes("-alpha", 1.0)
                self.main_hidden_for_setup = False
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception:
            pass

    def _has_mic_input_device(self) -> bool:
        if not MIC_AVAILABLE:
            return False
        try:
            devices = sd.query_devices()
            return any(d.get("max_input_channels", 0) > 0 for d in devices)
        except Exception:
            return False

    def _mic_input_candidates(self) -> list[tuple[int, int, str]]:
        candidates = []
        seen = set()
        devices = sd.query_devices()

        # Prefer OS default input first.
        try:
            default_in = sd.default.device[0]
            if isinstance(default_in, int) and default_in >= 0:
                d = devices[default_in]
                if d.get("max_input_channels", 0) > 0:
                    rate = int(d.get("default_samplerate") or MIC_SAMPLERATE)
                    candidates.append((default_in, max(8000, rate), d.get("name", f"Device {default_in}")))
                    seen.add(default_in)
        except Exception:
            pass

        # Then try every other input device.
        for idx, d in enumerate(devices):
            if idx in seen:
                continue
            if d.get("max_input_channels", 0) > 0:
                rate = int(d.get("default_samplerate") or MIC_SAMPLERATE)
                candidates.append((idx, max(8000, rate), d.get("name", f"Device {idx}")))

        return candidates

    def _start_mic_stream(self) -> bool:
        self._last_mic_device = "None"
        self._last_mic_error = ""

        if not MIC_AVAILABLE:
            self._last_mic_error = "sounddevice not installed"
            return False

        try:
            candidates = self._mic_input_candidates()
        except Exception as e:
            self._last_mic_error = f"Cannot enumerate mic devices: {e}"
            return False

        if not candidates:
            self._last_mic_error = "No input devices found in PortAudio"
            return False

        errors = []
        for device_idx, sample_rate, name in candidates:
            try:
                stream = sd.InputStream(
                    device=device_idx,
                    samplerate=sample_rate,
                    channels=1,
                    blocksize=MIC_CHUNK,
                    dtype="float32",
                    callback=self._mic_callback,
                )
                stream.start()
                self._audio_stream = stream
                self._last_mic_device = name
                print(f"[Mic] Using input device: {name} @ {sample_rate}Hz")
                return True
            except Exception as e:
                errors.append(f"{name}: {e}")

        self._last_mic_error = "; ".join(errors[:2]) if errors else "Unknown mic open error"
        return False

    def _detect_and_badge(self):
        self.has_mic = self._has_mic_input_device()
        self.has_mouse = PYNPUT_AVAILABLE
        self.has_winrt = self._has_winrt()
        self.has_aps = self._has_aps()

        self._set_badge(self.badge_mic, self.has_mic)
        self._set_badge(self.badge_mouse, self.has_mouse)
        self._set_badge(self.badge_accel, self.has_winrt or self.has_aps)

        accel_parts = []
        accel_parts.append("WinRT: yes" if self.has_winrt else "WinRT: no")
        accel_parts.append("APS: yes" if self.has_aps else "APS: no")
        reason = ""
        if not self.has_winrt and not self.has_aps:
            reason = " | Windows is not exposing an accelerometer backend (driver/hardware unavailable)."
        self.accel_info_var.set("Accel backend availability -> " + " | ".join(accel_parts) + reason)

        self._update_sounds_label()

    def _has_winrt(self) -> bool:
        if WINRT_ACCEL is None:
            return False
        try:
            return WINRT_ACCEL.get_default() is not None
        except Exception:
            return False

    def _open_aps_handle(self):
        kernel32 = ctypes.windll.kernel32
        # Some systems only work when write access is requested together with read.
        access_modes = [GENERIC_READ | GENERIC_WRITE, GENERIC_READ]

        for access in access_modes:
            handle = kernel32.CreateFileW(
                r"\\.\ShockMgr",
                access,
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                None,
                OPEN_EXISTING,
                0,
                None,
            )
            if handle != INVALID_HANDLE_VALUE:
                return handle

        return None

    def _has_aps(self) -> bool:
        handle = self._open_aps_handle()
        if handle is None:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True

    def _update_sounds_label(self):
        n = len(self.moan_sounds)
        src = f"{MOANS_DIR.name}/" if MOANS_DIR.exists() else "synth fallback"
        self.sounds_var.set(f"{n} sound{'s' if n != 1 else ''} loaded ({src})")

    def _update_moan_combo(self):
        if self.sound_names:
            values = ["Random"] + self.sound_names
            self.current_moan_combo["values"] = values
            if self.current_moan_var.get() not in values:
                self.current_moan_var.set("Random")
        else:
            self.current_moan_combo["values"] = ["No sounds"]
            self.current_moan_var.set("No sounds")

    def _reload_sounds(self):
        self.moan_sounds, self.sound_names = load_moan_sounds()
        self._update_moan_combo()
        self._update_sounds_label()

    def _mode_enabled(self) -> tuple[bool, bool, bool]:
        mode = self.detector_mode_var.get()
        if mode == "Mic only":
            return True, False, False
        if mode == "Accelerometer only":
            return False, True, False
        if mode == "Mouse jitter only":
            return False, False, True
        if mode == "Mic + accelerometer":
            return True, True, False
        if mode == "Mic + mouse jitter":
            return True, False, True
        return True, True, True

    def _on_detector_mode_change(self, _event=None):
        if self.running:
            self._set_status("Restarting sensors with new detector mode...", "#f5c542")
            self.stop()
            self.start()

    def _on_accel_backend_change(self, _event=None):
        if self.running:
            self._set_status("Restarting sensors with new accel backend...", "#f5c542")
            self.stop()
            self.start()

    def toggle(self):
        (self.stop if self.running else self.start)()

    def start(self):
        self._detect_and_badge()

        use_mic, use_accel, use_mouse = self._mode_enabled()
        backend_pref = self.accel_backend_var.get()

        started_any = False
        self.running = True
        mic_started = False

        if use_mic and not self.has_mic:
            self._last_mic_error = "No input device detected"

        if use_mic and self.has_mic:
            mic_started = self._start_mic_stream()
            if mic_started:
                started_any = True
            else:
                print(f"[Mic] {self._last_mic_error}")

        accel_started = False
        if use_accel:
            if backend_pref in ("Auto", "WinRT") and self.has_winrt:
                try:
                    accel = WINRT_ACCEL.get_default()
                    if accel is not None:
                        accel.report_interval = max(50, accel.minimum_report_interval)
                        self.accel_token = accel.add_reading_changed(self._accel_callback)
                        self.accel_obj = accel
                        self._last_accel_backend = "WinRT"
                        accel_started = True
                        started_any = True
                except Exception as e:
                    print(f"[Accel WinRT] {e}")

            if (not accel_started) and backend_pref in ("Auto", "ThinkPad APS"):
                try:
                    aps_h = self._open_aps_handle()
                    if aps_h is not None:
                        self.aps_handle = aps_h
                        self._aps_poller_thread = threading.Thread(
                            target=self._aps_poller,
                            daemon=True,
                        )
                        self._aps_poller_thread.start()
                        self._last_accel_backend = "ThinkPad APS"
                        accel_started = True
                        started_any = True
                        print("[Accel] ThinkPad APS activated")
                except Exception as e:
                    print(f"[APS] {e}")

            if not accel_started:
                self._last_accel_backend = "Unavailable"

        if use_mouse and self.has_mouse:
            self._prev_mouse_pos = None
            self._mouse_listener = pynput_mouse.Listener(on_move=self._mouse_move)
            self._mouse_listener.start()
            started_any = True

        if not started_any:
            self.running = False
            self._set_status(
                "No selected detector is available.\n"
                "Try Detector Mode: Auto, then Refresh sensor detection.",
                "#f09000",
            )
            return

        self.start_btn.config(text="STOP", bg="#333333")

        accel_hint = ""
        if use_accel and not accel_started:
            accel_hint = "\nAccel not found. On T470s, try backend: ThinkPad APS."
        mic_hint = ""
        if use_mic and not mic_started:
            mic_hint = (
                "\nMic failed to open. Check Windows mic privacy and device use. "
                f"Last error: {self._last_mic_error or 'unknown'}"
            )

        mic_label = self._last_mic_device if mic_started else "Unavailable"

        self._set_status(
            (
                f"Listening... slap it!\n"
                f"Mode: {self.detector_mode_var.get()} | Mic: {mic_label} | Accel: {self._last_accel_backend}"
                f"{accel_hint}{mic_hint}"
            ),
            "#4dff91",
        )

    def stop(self):
        self.running = False

        if self._audio_stream:
            try:
                self._audio_stream.stop()
                self._audio_stream.close()
            except Exception:
                pass
            self._audio_stream = None

        if self.accel_obj and self.accel_token is not None:
            try:
                self.accel_obj.remove_reading_changed(self.accel_token)
            except Exception:
                pass
            self.accel_obj = None
            self.accel_token = None

        if self.aps_handle:
            try:
                ctypes.windll.kernel32.CloseHandle(self.aps_handle)
            except Exception:
                pass
            self.aps_handle = None

        if self._mouse_listener:
            try:
                self._mouse_listener.stop()
            except Exception:
                pass
            self._mouse_listener = None

        self.start_btn.config(text="START", bg="#ff4d6d")
        self._set_status("Idle", "#666")
        self.root.after(0, self.impact_var.set, 0)
        self.root.after(0, self.impact_lbl.config, {"text": "-"})

    def _aps_poller(self):
        if not self.aps_handle:
            return

        kernel32 = ctypes.windll.kernel32
        prev_x = 0
        prev_y = 0

        while self.running and self.aps_handle is not None:
            buf = (ctypes.c_byte * 36)()
            bytes_returned = DWORD(0)

            success = kernel32.DeviceIoControl(
                self.aps_handle,
                IOCTL_APS_READ,
                None,
                0,
                byref(buf),
                36,
                byref(bytes_returned),
                None,
            )

            if success and bytes_returned.value >= 36:
                try:
                    data = bytes(buf[:36])
                    shorts = struct.unpack_from("<16h", data, 4)
                    x = shorts[1]
                    y = shorts[0]

                    if prev_x or prev_y:
                        dx = x - prev_x
                        dy = y - prev_y
                        delta_mag = math.sqrt(dx * dx + dy * dy)

                        pct = min(delta_mag / 200 * 100, 100)
                        self.root.after(0, self._update_meter, pct, f"APS delta {delta_mag:.0f}")

                        aps_thresh = self.accel_thresh.get() * 25.0
                        if delta_mag > aps_thresh:
                            intensity = min((delta_mag - aps_thresh) / 120.0, 1.0)
                            self._trigger(intensity)

                    prev_x = x
                    prev_y = y
                except Exception:
                    pass

            time.sleep(0.05)

    def _mic_callback(self, indata, _frames, _t, _status):
        if not self.running:
            return

        rms = float(np.sqrt(np.mean(indata ** 2)))
        pct = min(rms / 0.8 * 100, 100)
        self.root.after(0, self._update_meter, pct, f"Mic {rms:.3f} rms")

        if rms > self.mic_thresh.get():
            self._trigger(min(rms / 0.6, 1.0))

    def _accel_callback(self, _sender, args):
        if not self.running:
            return

        r = args.reading
        mag = math.sqrt(r.acceleration_x ** 2 + r.acceleration_y ** 2 + r.acceleration_z ** 2)
        pct = min(mag / 6.0 * 100, 100)
        self.root.after(0, self._update_meter, pct, f"Accel {mag:.2f} g")

        if mag > self.accel_thresh.get():
            self._trigger(min((mag - self.accel_thresh.get()) / 3.0, 1.0))

    def _mouse_move(self, x, y):
        if not self.running:
            return

        now = time.time()
        if self._prev_mouse_pos:
            dx = x - self._prev_mouse_pos[0]
            dy = y - self._prev_mouse_pos[1]
            dt = now - self._prev_mouse_t

            if 0 < dt < 0.05:
                speed = math.sqrt(dx ** 2 + dy ** 2) / dt
                pct = min(speed / 5000 * 100, 100)
                self.root.after(0, self._update_meter, pct, f"Mouse {speed:.0f} px/s")
                if speed > MOUSE_SPEED_MIN:
                    self._trigger(min(speed / 8000, 1.0))

        self._prev_mouse_pos = (x, y)
        self._prev_mouse_t = now

    def _trigger(self, intensity: float = 0.5):
        now = time.time()
        if (now - self.last_moan) < COOLDOWN_SEC:
            return

        self.last_moan = now
        threading.Thread(target=self._play_moan, args=(intensity,), daemon=True).start()

    def _play_moan(self, _intensity: float):
        if not self.moan_sounds:
            return

        selected = self.current_moan_var.get()
        if selected == "Random":
            sound = random.choice(self.moan_sounds)
        else:
            try:
                idx = self.sound_names.index(selected)
                sound = self.moan_sounds[idx]
            except ValueError:
                sound = random.choice(self.moan_sounds)

        sound.play()
        self.moan_count += 1
        self.root.after(0, self.count_var.set, f"Moans: {self.moan_count}")
        self.root.after(0, self._set_status, "SLAPPED", "#ff4d6d")
        self.root.after(
            900,
            lambda: self.running and self._set_status(
                (
                    f"Listening... slap it!\n"
                    f"Mode: {self.detector_mode_var.get()} | "
                    f"Mic: {self._last_mic_device if self._audio_stream else 'Unavailable'} | "
                    f"Accel: {self._last_accel_backend}"
                ),
                "#4dff91",
            ),
        )

    def _test_moan(self):
        threading.Thread(target=self._play_moan, args=(0.6,), daemon=True).start()

    def _update_meter(self, pct, label):
        self.impact_var.set(pct)
        self.impact_lbl.config(text=label)

    def _set_status(self, msg, color="#666"):
        self.status_var.set(msg)
        self.status_lbl.config(fg=color)

    def cleanup(self):
        try:
            self.scroll_canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        self._close_dependency_modal()
        self.stop()
        if PYGAME_AVAILABLE:
            try:
                pygame.mixer.quit()
            except Exception:
                pass


if __name__ == "__main__":
    root = tk.Tk()
    app = SlapMoanApp(root)
    root.protocol("WM_DELETE_WINDOW", lambda: (app.cleanup(), root.destroy()))
    root.mainloop()
