<div align="center">

```
███████╗██╗      █████╗ ██████╗ ███╗   ██╗███╗   ███╗ ██████╗  █████╗ ███╗   ██╗
██╔════╝██║     ██╔══██╗██╔══██╗████╗  ██║████╗ ████║██╔═══██╗██╔══██╗████╗  ██║
███████╗██║     ███████║██████╔╝██╔██╗ ██║██╔████╔██║██║   ██║███████║██╔██╗ ██║
╚════██║██║     ██╔══██║██╔═══╝ ██║╚██╗██║██║╚██╔╝██║██║   ██║██╔══██║██║╚██╗██║
███████║███████╗██║  ██║██║     ██║ ╚████║██║ ╚═╝ ██║╚██████╔╝██║  ██║██║ ╚████║
╚══════╝╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝  ╚═══╝╚═╝     ╚═╝ ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═══╝
```

**Slap your laptop. Hear the consequences.**

[![Python](https://img.shields.io/badge/Python-3.9%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?style=flat-square&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Gloriously_Dumb-ff69b4?style=flat-square)]()

</div>

---

## What is this?

**SlapMoan** is a desktop app that turns your laptop into a dramatic entity with feelings. Hit it, shake it, or tap the case — and it *expresses itself*. Powered by your microphone, accelerometer, or mouse jitter, SlapMoan detects physical impact and responds with a random moan sound.

It's stupid. It's perfect.

---

## Features

| | Feature | Description |
|---|---|---|
| 🎤 | **Multi-sensor Detection** | Mic audio spikes, WinRT/APS accelerometer, or mouse jitter — pick your poison |
| 🔊 | **Custom Sound Library** | Drop `.mp3`, `.wav`, or `.ogg` files into `moans/` and they're in the rotation |
| 🌑 | **Dark UI** | Clean Tkinter interface with sensor badges, impact meter, and sound controls |
| 🔧 | **Dependency Helper** | Built-in checker and installer — no hunting around for missing packages |
| 🎛️ | **Flexible Modes** | Mic only, accelerometer only, mouse only, or stack them together |

---

## Installation

### 1 · Clone the repo

```sh
git clone https://github.com/ke1thdev/slapnmoan.git
cd slapmoan
```

### 2 · Set up your environment

Python **3.9+** is required. A venv is optional but recommended:

```sh
python -m venv .venv
.venv\Scripts\activate   # Windows
```

### 3 · Install dependencies

```sh
pip install -r requirements.txt
```

> **Accelerometer?** Check `requirements.txt` and uncomment the right backend for your hardware:
> - **WinRT** (Surface, most modern laptops) → `winrt-runtime` + `winrt-Windows.Devices.Sensors`
> - **ThinkPad APS** (older ThinkPads) → `winsdk`

---

## Usage

### Add your sounds *(optional)*

Sound files are already included in `moans/`. Swap them out or add your own — just name them numerically (`1.mp3`, `2.wav`, `3.ogg`, etc.) and drop them in the folder.

### Launch

```sh
python slap2.py
```

The UI shows your detected sensors. Now slap it.

### Controls

| Control | What it does |
|---|---|
| `START / STOP` | Toggle impact detection |
| `TEST MOAN` | Instant gratification — plays a random sound |
| `Detector Mode` | Switch between mic / accel / mouse / combined |
| `Reload Sounds` | Picks up any new files you added to `moans/` |

---

## Dependencies

```
numpy
pygame
sounddevice        # microphone detection
pynput             # mouse jitter fallback
winrt-runtime      # WinRT accelerometer (Surface / modern laptops)
winrt-Windows.Devices.Sensors
winsdk             # ThinkPad APS (alternative, see requirements.txt)
```

All installable via `pip`. See `requirements.txt` for version details and hardware-specific comments.

---

## Troubleshooting

<details>
<summary><b>🔇 No sound playing</b></summary>

Make sure `moans/` has at least one valid audio file. If it's empty, the app falls back to a synthesized beep tone instead.

</details>

<details>
<summary><b>📡 No sensors detected</b></summary>

- **Mic**: Confirm your microphone is enabled in Windows Settings → Privacy → Microphone, and isn't currently claimed by another app.
- **Accelerometer**: Your device needs a supported sensor *and* the matching backend installed. Check `requirements.txt` and install the right one for your hardware.

</details>

<details>
<summary><b>🔒 Permission errors</b></summary>

Windows may block mic access depending on your privacy settings. Go to **Settings → Privacy & Security → Microphone** and enable access for desktop apps.

</details>

---

## Contributing

Pull requests are welcome. Open an issue for bugs, feature requests, or hardware compatibility reports. If your laptop makes interesting noises when slapped, that's probably a separate issue.

---

## License

[MIT](LICENSE) — do whatever you want with it.

---

<div align="center">

*Made with poor judgment and excellent results.*

</div>