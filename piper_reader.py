"""
Piper TTS Document Reader
Reads .docx, .md, and .txt files aloud with full playback controls.
"""
import os
import re
import sys
import shutil
import wave
import time
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pygame
from docx import Document
from piper import PiperVoice, SynthesisConfig

# Voice model filename (config is this + ".json")
MODEL_NAME = "en_US-lessac-medium.onnx"


def resolve_model() -> str:
    """Locate the voice model.

    Search order: the PyInstaller bundle (model shipped inside the .exe),
    the directory next to the executable/script, then the user's home dir.
    Returns the first path whose .onnx AND .onnx.json both exist; if none
    match, returns the home-dir path so the load error is clear.
    """
    candidates = []
    # PyInstaller onefile unpacks --add-data files into sys._MEIPASS
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(os.path.join(meipass, MODEL_NAME))
    # Next to the frozen exe (or the script when run from source)
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(base, MODEL_NAME))
    # User home dir (user-supplied / legacy location)
    home_path = os.path.join(os.path.expanduser("~"), MODEL_NAME)
    candidates.append(home_path)

    for path in candidates:
        if os.path.exists(path) and os.path.exists(path + ".json"):
            return path
    return home_path


DEFAULT_MODEL = resolve_model()

# Speed presets: label -> length_scale (higher = slower)
SPEED_PRESETS = {
    "2.0x": 0.5,
    "1.5x": 0.67,
    "1.25x": 0.8,
    "1.0x": 1.0,
    "0.75x": 1.33,
    "0.5x": 2.0,
}


def clean_text(text: str) -> str:
    """Clean document text for TTS consumption."""
    # Smart quotes -> straight
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    # Dashes -> hyphen
    text = text.replace("\u2013", "-").replace("\u2014", "-")
    # Ellipsis
    text = text.replace("\u2026", "...")
    # Bullets and list markers
    text = re.sub(r"[\u2022\u2023\u25e6\u2043\u2219\u25aa\u25ab\u25cf\u25cb\u25a0\u25a1\u25b6\u25c6\u00b7]", "", text)
    # Non-breaking space / narrow no-break space
    text = text.replace("\u00a0", " ").replace("\u202f", " ")
    # Zero-width chars
    text = re.sub(r"[\u200b\u200c\u200d\ufeff]", "", text)
    # Soft hyphen
    text = text.replace("\u00ad", "")
    # Section / pilcrow / degree symbols spoken oddly
    text = text.replace("\u00a7", "section ").replace("\u00b6", "")
    # Copyright, trademark, registered
    text = text.replace("\u00a9", "copyright ").replace("\u2122", "trademark ").replace("\u00ae", "registered ")
    # Curly apostrophes that might remain
    text = text.replace("\u02bc", "'")
    # Arrows -> words
    text = text.replace("\u2192", " to ").replace("\u2190", " from ")
    text = text.replace("\u2194", " ").replace("\u21d2", " implies ")
    # Strip emojis and misc symbols (geometric shapes, dingbats, emoticons, etc.)
    text = re.sub(r"[\u2500-\u27bf\u2b00-\u2bff\U0001f300-\U0001f9ff\U0001fa00-\U0001faff]", "", text)
    # Horizontal rules / decorative separators (lines of dashes, underscores, equals, asterisks)
    text = re.sub(r"^[\-=_*~#]{3,}\s*$", "", text, flags=re.MULTILINE)
    # Standalone symbols on their own or clustered (e.g. "* # / etc" or "***")
    # Remove symbols that aren't part of words — keep hyphens in hyphenated words
    text = re.sub(r"(?<!\w)[*#/\\|~^@&<>{}[\]()=_+]+(?!\w)", " ", text)
    # Markdown-style bold/italic markers around words: **word** or *word*
    text = re.sub(r"\*{1,3}([^*]+?)\*{1,3}", r"\1", text)
    # Any remaining lone asterisks
    text = text.replace("*", "")
    # Hash marks used as headers: "## Title" -> "Title"
    text = re.sub(r"^#+\s*", "", text, flags=re.MULTILINE)
    # Strip remaining control chars (keep newlines, tabs)
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    # Collapse multiple spaces
    text = re.sub(r" {2,}", " ", text)
    # Collapse excessive newlines (3+ -> 2)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def clean_markdown(md: str) -> str:
    """Strip Markdown syntax so the text reads naturally aloud."""
    # Fenced code blocks (``` and ~~~) — drop entirely
    md = re.sub(r"```.*?```", "", md, flags=re.DOTALL)
    md = re.sub(r"~~~.*?~~~", "", md, flags=re.DOTALL)
    # HTML comments and tags (badges, <br>, <div>, etc.)
    md = re.sub(r"<!--.*?-->", "", md, flags=re.DOTALL)
    md = re.sub(r"<[^>]+>", "", md)
    # Images / badges: ![alt](url) and ![alt][ref] -> drop
    md = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", md)
    md = re.sub(r"!\[[^\]]*\]\[[^\]]*\]", "", md)
    # Links: [text](url) -> text, [text][ref] -> text
    md = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", md)
    md = re.sub(r"\[([^\]]+)\]\[[^\]]*\]", r"\1", md)
    # Link/badge reference definitions:  [ref]: https://...
    md = re.sub(r"^\s*\[[^\]]+\]:\s*\S+.*$", "", md, flags=re.MULTILINE)
    # Inline code `code` -> code
    md = re.sub(r"`+([^`]+)`+", r"\1", md)
    # ATX headings: "## Title" -> "Title"
    md = re.sub(r"^\s{0,3}#{1,6}\s*", "", md, flags=re.MULTILINE)
    # Setext heading underlines and horizontal rules
    md = re.sub(r"^\s*[=\-_*]{3,}\s*$", "", md, flags=re.MULTILINE)
    # Blockquote markers
    md = re.sub(r"^\s*>+\s?", "", md, flags=re.MULTILINE)
    # Bullet and numbered list markers
    md = re.sub(r"^\s*[-*+]\s+", "", md, flags=re.MULTILINE)
    md = re.sub(r"^\s*\d+\.\s+", "", md, flags=re.MULTILINE)
    # Table separator rows (|---|:--:|) -> drop
    md = re.sub(r"^\s*\|?[\s:\-|]{3,}\|?\s*$", "", md, flags=re.MULTILINE)
    # Remaining table pipes -> spaces
    md = re.sub(r"\s*\|\s*", " ", md)
    # Strikethrough ~~text~~ -> text
    md = re.sub(r"~~(.+?)~~", r"\1", md)
    # clean_text handles **bold**/*italic*, stray symbols, etc.
    return clean_text(md)


def extract_markdown_text(path: str) -> str:
    """Read a Markdown file and strip syntax for TTS."""
    with open(path, encoding="utf-8", errors="replace") as f:
        return clean_markdown(f.read())


def extract_txt_text(path: str) -> str:
    """Read a plain-text file and clean it for TTS."""
    with open(path, encoding="utf-8", errors="replace") as f:
        return clean_text(f.read())


def extract_docx_text(path: str) -> str:
    """Extract and clean text from a .docx file."""
    doc = Document(path)
    paragraphs = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            paragraphs.append(t)
    raw = "\n\n".join(paragraphs)
    return clean_text(raw)


def extract_text(path: str) -> str:
    """Dispatch to the right extractor based on file extension."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return extract_docx_text(path)
    if ext in (".md", ".markdown"):
        return extract_markdown_text(path)
    # Everything else is treated as plain text (.txt, .text, unknown)
    return extract_txt_text(path)


def get_wav_duration(path: str) -> float:
    """Get duration of a WAV file in seconds."""
    with wave.open(path, "rb") as wf:
        frames = wf.getnframes()
        rate = wf.getframerate()
        return frames / rate if rate else 0.0


def fmt_time(seconds: float) -> str:
    """Format seconds as MM:SS."""
    m, s = divmod(max(0, int(seconds)), 60)
    return f"{m:02d}:{s:02d}"


class PiperReaderApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Piper Document Reader")
        self.root.geometry("700x550")
        self.root.minsize(600, 450)

        self.wav_path = None
        self.duration = 0.0
        self.playing = False
        self.paused = False
        self.play_start_time = 0.0  # wall-clock when play/resume started
        self.play_offset = 0.0     # accumulated offset in seconds
        self.seeking = False
        self.voice = None
        self.generating = False

        pygame.mixer.init()

        self._build_ui()
        self._update_loop()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 4}

        # --- File selection ---
        file_frame = ttk.Frame(self.root)
        file_frame.pack(fill="x", **pad)

        ttk.Label(file_frame, text="Document:").pack(side="left")
        self.file_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_var)
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(4, 4))
        ttk.Button(file_frame, text="Browse", command=self._browse).pack(side="left")
        ttk.Button(file_frame, text="Load", command=self._load_doc).pack(side="left", padx=(4, 0))

        # --- Text preview ---
        preview_frame = ttk.LabelFrame(self.root, text="Text Preview")
        preview_frame.pack(fill="both", expand=True, **pad)

        self.text_area = tk.Text(preview_frame, wrap="word", state="disabled",
                                font=("Consolas", 10))
        scrollbar = ttk.Scrollbar(preview_frame, command=self.text_area.yview)
        self.text_area.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.text_area.pack(fill="both", expand=True)

        # --- Generate button + speed ---
        gen_frame = ttk.Frame(self.root)
        gen_frame.pack(fill="x", **pad)

        self.gen_btn = ttk.Button(gen_frame, text="Generate Audio", command=self._generate)
        self.gen_btn.pack(side="left")

        ttk.Label(gen_frame, text="Speed:").pack(side="left", padx=(16, 4))
        self.speed_var = tk.StringVar(value="1.0x")
        speed_combo = ttk.Combobox(gen_frame, textvariable=self.speed_var,
                                   values=list(SPEED_PRESETS.keys()),
                                   state="readonly", width=6)
        speed_combo.pack(side="left")

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(gen_frame, variable=self.progress_var,
                                            maximum=100, length=200)
        self.progress_bar.pack(side="left", padx=(16, 0), fill="x", expand=True)

        self.status_label = ttk.Label(gen_frame, text="")
        self.status_label.pack(side="right", padx=(8, 0))

        # --- Playback controls ---
        ctrl_frame = ttk.Frame(self.root)
        ctrl_frame.pack(fill="x", **pad)

        self.rew_btn = ttk.Button(ctrl_frame, text="<< 10s", command=self._rewind, state="disabled")
        self.rew_btn.pack(side="left")

        self.play_btn = ttk.Button(ctrl_frame, text="Play", command=self._play_pause, state="disabled")
        self.play_btn.pack(side="left", padx=4)

        self.fwd_btn = ttk.Button(ctrl_frame, text="10s >>", command=self._forward, state="disabled")
        self.fwd_btn.pack(side="left")

        self.stop_btn = ttk.Button(ctrl_frame, text="Stop", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", padx=4)

        self.save_btn = ttk.Button(ctrl_frame, text="Save WAV", command=self._save_wav, state="disabled")
        self.save_btn.pack(side="left", padx=4)
        self.stop_btn.pack(side="left", padx=4)

        # --- Timeline ---
        time_frame = ttk.Frame(self.root)
        time_frame.pack(fill="x", **pad)

        self.time_label = ttk.Label(time_frame, text="00:00")
        self.time_label.pack(side="left")

        self.slider = ttk.Scale(time_frame, from_=0, to=100, orient="horizontal",
                                command=self._on_slider)
        self.slider.pack(side="left", fill="x", expand=True, padx=8)
        self.slider.bind("<ButtonPress-1>", lambda e: setattr(self, "seeking", True))
        self.slider.bind("<ButtonRelease-1>", self._on_slider_release)

        self.dur_label = ttk.Label(time_frame, text="00:00")
        self.dur_label.pack(side="right")

    def _browse(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Supported Documents", "*.docx *.md *.markdown *.txt *.text"),
                ("Word Documents", "*.docx"),
                ("Markdown", "*.md *.markdown"),
                ("Text", "*.txt *.text"),
                ("All files", "*.*"),
            ]
        )
        if path:
            self.file_var.set(path)
            self._load_doc()

    def _load_doc(self):
        path = self.file_var.get().strip()
        if not path or not os.path.isfile(path):
            messagebox.showerror("Error", "File not found.")
            return
        try:
            text = extract_text(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read document:\n{e}")
            return

        self.text_area.configure(state="normal")
        self.text_area.delete("1.0", "end")
        self.text_area.insert("1.0", text)
        self.text_area.configure(state="disabled")
        self.status_label.configure(text=f"{len(text)} chars loaded")

    def _get_text(self) -> str:
        return self.text_area.get("1.0", "end").strip()

    def _generate(self):
        text = self._get_text()
        if not text:
            messagebox.showwarning("Warning", "No text to generate. Load a document first.")
            return
        if self.generating:
            return

        self._stop()
        self.generating = True
        self.gen_btn.configure(state="disabled")
        self.status_label.configure(text="Generating...")
        self.progress_var.set(0)

        threading.Thread(target=self._generate_worker, args=(text,), daemon=True).start()

    def _generate_worker(self, text: str):
        try:
            if self.voice is None:
                if not (os.path.exists(DEFAULT_MODEL) and os.path.exists(DEFAULT_MODEL + ".json")):
                    raise FileNotFoundError(
                        f"Voice model not found.\n\nExpected '{MODEL_NAME}' (and "
                        f"'{MODEL_NAME}.json') bundled with the app, next to the "
                        f"executable, or in your home folder:\n{os.path.expanduser('~')}\n\n"
                        "See the README for where to download a Piper voice."
                    )
                self.root.after(0, lambda: self.status_label.configure(text="Loading voice model..."))
                self.voice = PiperVoice.load(DEFAULT_MODEL)

            speed_label = self.speed_var.get()
            length_scale = SPEED_PRESETS.get(speed_label, 1.0)
            syn_config = SynthesisConfig(length_scale=length_scale)

            # Generate to temp wav
            tmp = tempfile.NamedTemporaryFile(suffix=".wav", delete=False, dir=tempfile.gettempdir())
            tmp_path = tmp.name
            tmp.close()

            # Split text into chunks for progress reporting
            sentences = re.split(r'(?<=[.!?])\s+', text)
            sentences = [s for s in sentences if s.strip()]
            total = len(sentences)

            with wave.open(tmp_path, "wb") as wf:
                for i, sentence in enumerate(sentences):
                    self.voice.synthesize_wav(sentence, wf, syn_config=syn_config,
                                              set_wav_format=(i == 0))
                    pct = ((i + 1) / total) * 100
                    self.root.after(0, lambda p=pct: self.progress_var.set(p))

            # Clean up old temp file
            if self.wav_path and os.path.exists(self.wav_path):
                try:
                    os.unlink(self.wav_path)
                except OSError:
                    pass

            self.wav_path = tmp_path
            self.duration = get_wav_duration(tmp_path)

            self.root.after(0, self._on_generate_done)

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", f"Generation failed:\n{e}"))
            self.root.after(0, lambda: self.status_label.configure(text="Error"))
        finally:
            self.generating = False
            self.root.after(0, lambda: self.gen_btn.configure(state="normal"))

    def _on_generate_done(self):
        self.progress_var.set(100)
        self.status_label.configure(text=f"Ready ({fmt_time(self.duration)})")
        self.slider.configure(to=self.duration)
        self.dur_label.configure(text=fmt_time(self.duration))
        self.play_btn.configure(state="normal")
        self.rew_btn.configure(state="normal")
        self.fwd_btn.configure(state="normal")
        self.stop_btn.configure(state="normal")
        self.save_btn.configure(state="normal")

    def _save_wav(self):
        if not self.wav_path or not os.path.exists(self.wav_path):
            messagebox.showwarning("Warning", "No audio to save. Generate first.")
            return
        # Default filename from the source doc name
        src = self.file_var.get().strip()
        default_name = os.path.splitext(os.path.basename(src))[0] + ".wav" if src else "output.wav"
        dest = filedialog.asksaveasfilename(
            defaultextension=".wav",
            initialfile=default_name,
            filetypes=[("WAV Audio", "*.wav"), ("All files", "*.*")]
        )
        if dest:
            shutil.copy2(self.wav_path, dest)
            self.status_label.configure(text=f"Saved: {os.path.basename(dest)}")

    def _play_pause(self):
        if not self.wav_path:
            return

        if not self.playing:
            # Start playback
            pygame.mixer.music.load(self.wav_path)
            pygame.mixer.music.play(start=self.play_offset)
            self.play_start_time = time.time()
            self.playing = True
            self.paused = False
            self.play_btn.configure(text="Pause")
        elif self.paused:
            # Resume
            pygame.mixer.music.unpause()
            self.play_start_time = time.time()
            self.paused = False
            self.play_btn.configure(text="Pause")
        else:
            # Pause
            self.play_offset = self._current_pos()
            pygame.mixer.music.pause()
            self.paused = True
            self.play_btn.configure(text="Play")

    def _stop(self):
        if self.playing:
            pygame.mixer.music.stop()
        self.playing = False
        self.paused = False
        self.play_offset = 0.0
        self.play_btn.configure(text="Play")
        self.slider.set(0)
        self.time_label.configure(text="00:00")

    def _current_pos(self) -> float:
        """Get current playback position in seconds."""
        if not self.playing:
            return self.play_offset
        if self.paused:
            return self.play_offset
        elapsed = time.time() - self.play_start_time
        return min(self.play_offset + elapsed, self.duration)

    def _seek_to(self, seconds: float):
        """Seek to a position in seconds."""
        seconds = max(0.0, min(seconds, self.duration))
        was_playing = self.playing and not self.paused

        if self.playing:
            pygame.mixer.music.stop()

        self.play_offset = seconds
        self.playing = False
        self.paused = False

        if was_playing:
            pygame.mixer.music.load(self.wav_path)
            pygame.mixer.music.play(start=seconds)
            self.play_start_time = time.time()
            self.playing = True
            self.play_btn.configure(text="Pause")
        else:
            self.play_btn.configure(text="Play")

    def _rewind(self):
        pos = self._current_pos()
        self._seek_to(pos - 10)

    def _forward(self):
        pos = self._current_pos()
        self._seek_to(pos + 10)

    def _on_slider(self, value):
        if self.seeking:
            self.time_label.configure(text=fmt_time(float(value)))

    def _on_slider_release(self, event):
        self.seeking = False
        val = self.slider.get()
        self._seek_to(float(val))

    def _update_loop(self):
        """Update slider and time label every 100ms."""
        if self.playing and not self.paused and not self.seeking:
            pos = self._current_pos()
            if pos >= self.duration:
                self._stop()
            else:
                self.slider.set(pos)
                self.time_label.configure(text=fmt_time(pos))

        self.root.after(100, self._update_loop)


def main():
    root = tk.Tk()
    app = PiperReaderApp(root)

    # If a file path was passed as argument, load it
    if len(sys.argv) > 1:
        path = sys.argv[1]
        app.file_var.set(path)
        app.root.after(100, app._load_doc)

    root.mainloop()

    # Cleanup
    pygame.mixer.quit()
    if app.wav_path and os.path.exists(app.wav_path):
        try:
            os.unlink(app.wav_path)
        except OSError:
            pass


if __name__ == "__main__":
    main()
