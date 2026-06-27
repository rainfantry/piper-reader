# Piper Document Reader

**A simple desktop app that reads your documents aloud using [Piper](https://github.com/OHF-Voice/piper1-gpl) neural text-to-speech.**

Load a Word document, Markdown file, or plain-text file, generate natural-sounding speech locally on your machine, and listen with full playback controls — play/pause, skip, seek, adjustable speed, and save to WAV. No cloud, no API keys, everything runs offline.

![Python](https://img.shields.io/badge/python-3.10%2B-blue) ![Platform](https://img.shields.io/badge/platform-Windows-lightgrey) ![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

- **Reads `.docx`, `.md`/`.markdown`, and `.txt` files** — Markdown syntax (headings, badges, links, tables, code spans) is stripped automatically so it reads naturally.
- **Offline neural TTS** — powered by Piper voice models; no internet connection required.
- **Full playback controls** — play, pause, stop, rewind 10s, forward 10s, and a draggable seek bar.
- **Adjustable speed** — 0.5x to 2.0x presets.
- **Save to WAV** — export the generated audio to a file.
- **Drag-and-drop friendly CLI** — pass a file path as an argument to load it on launch.

## Quick Start

### Option 1 — Download the executable

Grab the latest `PiperReader.exe` from the [Releases](https://github.com/rainfantry/piper-reader/releases) page. No Python install needed.

You still need a Piper voice model (see [Voice model](#voice-model) below).

### Option 2 — Run from source

```bash
git clone https://github.com/rainfantry/piper-reader.git
cd piper-reader
pip install -r requirements.txt
python piper_reader.py
```

## Voice model

The app expects a Piper voice model in your home folder by default:

```
~/en_US-lessac-medium.onnx
~/en_US-lessac-medium.onnx.json
```

Download voices from the [Piper voices collection](https://huggingface.co/rhasspy/piper-voices). Place both the `.onnx` and `.onnx.json` files side by side. To use a different model, edit the `DEFAULT_MODEL` path near the top of `piper_reader.py`.

## Usage

1. Launch the app (`PiperReader.exe` or `python piper_reader.py`).
2. Click **Browse** and pick a `.docx`, `.md`, or `.txt` file — or paste a path and click **Load**.
3. Review the extracted text in the preview pane.
4. Pick a **Speed**, then click **Generate Audio**.
5. Use the playback controls to listen, and **Save WAV** to export.

You can also launch with a file directly:

```bash
python piper_reader.py "C:\path\to\document.md"
```

## Requirements

- Python 3.10 or newer
- [`piper-tts`](https://pypi.org/project/piper-tts/)
- [`pygame`](https://pypi.org/project/pygame/) (audio playback)
- [`python-docx`](https://pypi.org/project/python-docx/) (Word document support)
- `tkinter` (bundled with standard CPython on Windows)

## Building the executable

The release binary is built with [PyInstaller](https://pyinstaller.org/):

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name PiperReader piper_reader.py
```

The result lands in `dist/PiperReader.exe`.

## License

Released under the [MIT License](LICENSE).
