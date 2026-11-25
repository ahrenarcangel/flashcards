#!/usr/bin/env python3
"""
Flash Card Game - graphical version (Tkinter).

Uses the same Excel workbook as flashcards.py (keep both files in one folder).
  * "Cards"    sheet = your questions and correct answers (edit in Excel)
  * "Scores"   sheet = leaderboard history
  * "Settings" sheet = remembers your light/dark choice (created automatically)

Run:   python flashcards_gui.py [path/to/deck.xlsx]
Build: pyinstaller --onefile --noconsole flashcards_gui.py
"""

import array
import atexit
import math
import os
import random
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import wave
from pathlib import Path
from tkinter import font as tkfont
from tkinter import messagebox

from flashcards import (DEFAULT_FILE, Store, app_dir, is_correct,
                        make_hint, pick_questions)

FONT = "Inter"             # replaced at start-up; falls back to Segoe UI


def load_bundled_fonts():
    """Windows: register every .ttf/.otf in a 'fonts' folder for this program
    only (nothing is installed on the computer). Looks inside the PyInstaller
    bundle first, then next to the program."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        FR_PRIVATE = 0x10
        bases = {Path(getattr(sys, "_MEIPASS", app_dir())), app_dir()}
        for base in bases:
            for p in sorted((base / "fonts").glob("*")):
                if p.suffix.lower() in (".ttf", ".otf"):
                    ctypes.windll.gdi32.AddFontResourceExW(str(p),
                                                           FR_PRIVATE, 0)
    except Exception:
        pass


def pick_font(families):
    """Use Inter (or a variant like 'Inter 18pt') if available."""
    names = {f.lower(): f for f in families}
    if "inter" in names:
        return names["inter"]
    for low in sorted(names):
        if low.startswith("inter "):
            return names[low]
    return "Segoe UI"
ALL = "All categories"
MODE_KEYS = ("id", "mc", "tf", "cata")
MODE_NAMES = {"id": "Identification", "mc": "Multiple choice",
              "tf": "True or False", "cata": "Check all that apply"}
MODE_SHORT = {"id": "ID", "mc": "MC", "tf": "T/F", "cata": "CATA"}
MODE_TAG = {"id": "QUESTION", "mc": "QUESTION", "tf": "TRUE OR FALSE",
            "cata": "SELECT ALL THAT APPLY"}
MAX_CATA_OPTIONS = 6

THEMES = {
    "light": dict(
        bg="#f3f5fb", surface="#e4e9f4", card="#ffffff", card_border="#d7ddec",
        shadow="#d2d8e8", text="#1c2333", muted="#6b7385",
        accent="#3b6fd4", accent_hover="#2f5ab0", accent_text="#ffffff",
        good="#2e9e5b", good_bg="#e4f6ec", bad="#d64545", bad_bg="#fdeaea",
        warn="#e08a1e", neutral_bg="#e8effd",
        entry_bg="#ffffff", entry_border="#cdd5e6"),
    "dark": dict(
        bg="#12151c", surface="#252b3a", card="#1d222e", card_border="#2f3749",
        shadow="#0a0c11", text="#e8ecf5", muted="#8b94a9",
        accent="#6c9bff", accent_hover="#5a88ee", accent_text="#0f1420",
        good="#4cc383", good_bg="#16301f", bad="#ff6b6b", bad_bg="#3a1c20",
        warn="#f2a444", neutral_bg="#1f2a44",
        entry_bg="#1d222e", entry_border="#323a4e"),
}


# --------------------------------------------------------------------------
# Small helpers: colours, shapes, easing, animation
# --------------------------------------------------------------------------
def _rgb(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def mix(a, b, t):
    ra, rb = _rgb(a), _rgb(b)
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t)
                                   for x, y in zip(ra, rb))


def mix_theme(a, b, t):
    return {k: mix(a[k], b[k], t) for k in a}


def rr_points(x1, y1, x2, y2, r):
    """Points for a rounded rectangle; draw with create_polygon(smooth=True)."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
            x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]


def ease_in_out(p):
    return p * p * (3 - 2 * p)


def ease_out(p):
    return 1 - (1 - p) ** 3


def ease_in(p):
    return p ** 3


def animate(widget, ms, step, done=None, ease=ease_in_out):
    """Call step(p) with p going 0 -> 1 over `ms`. Returns a cancel function."""
    start = time.perf_counter()
    state = {"job": None, "stop": False}

    def tick():
        if state["stop"]:
            return
        p = min(1.0, (time.perf_counter() - start) * 1000.0 / ms)
        step(ease(p))
        if p < 1.0:
            state["job"] = widget.after(12, tick)
        elif done:
            done()

    def cancel():
        state["stop"] = True
        if state["job"] is not None:
            try:
                widget.after_cancel(state["job"])
            except Exception:
                pass

    tick()
    return cancel


def post_menu(widget, theme, options, on_pick):
    m = tk.Menu(widget, tearoff=0, bd=0, relief="flat", font=(FONT, 11),
                bg=theme["card"], fg=theme["text"],
                activebackground=theme["accent"],
                activeforeground=theme["accent_text"])
    for o in options:
        m.add_command(label=o, command=lambda o=o: on_pick(o))
    try:
        m.tk_popup(widget.winfo_rootx(),
                   widget.winfo_rooty() + widget.winfo_height() + 4)
    finally:
        m.grab_release()


# --------------------------------------------------------------------------
# Text fitting: pick a font size so wrapped text fits in a box
# --------------------------------------------------------------------------
_FONTS = {}


def _font(root, size, bold):
    key = (FONT, size, bold)
    f = _FONTS.get(key)
    if f is None:
        f = _FONTS[key] = tkfont.Font(root=root, family=FONT, size=size,
                                      weight="bold" if bold else "normal")
    return f


def wrap_height(root, text, size, bold, width):
    """Pixel height of `text` wrapped at `width` (simulates Tk's word wrap)."""
    f = _font(root, size, bold)
    space, lines = f.measure(" "), 0
    for para in str(text).split("\n"):
        cur, n = 0, 1
        for word in para.split(" "):
            wl = f.measure(word)
            if cur == 0:
                cur = wl
            elif cur + space + wl <= width:
                cur += space + wl
            else:
                n, cur = n + 1, wl
            if wl > width:                      # a single very long word
                n += int(wl // width)
        lines += n
    return lines * f.metrics("linespace")


def fit_text(root, text, size, bold, width, max_h, minimum=8):
    """Largest size (<= `size`) at which text fits; trims with '...' as a
    last resort. Returns (size, text)."""
    while size > minimum and wrap_height(root, text, size, bold, width) > max_h:
        size -= 1
    text = str(text)
    while len(text) > 4 and wrap_height(root, text, size, bold, width) > max_h:
        text = text[:max(4, int(len(text) * 0.85))].rstrip() + "\u2026"
    return size, text


# --------------------------------------------------------------------------
# Sound effects: tiny, quiet tones synthesised at start-up (no audio files)
# --------------------------------------------------------------------------
class Sounds:
    """Soft sine-based chimes and filtered-noise swishes.

    Everything is kept well below full volume (peaks of roughly 10-25%), with
    gentle fade-ins and fade-outs so nothing clicks or startles.
    Plays through winsound on Windows, afplay on macOS, paplay/aplay on Linux.
    """
    RATE = 22050

    def __init__(self, enabled=True):
        self.enabled = enabled
        self.ready, self.files, self.dir = False, {}, None
        self._loading, self._winsound, self._cmd = False, None, None

    # ---- synthesis -------------------------------------------------------
    def _tone(self, freq, dur, vol, attack=0.012, decay=7.0,
              harm=((1, 1.0), (2, 0.18), (3, 0.05))):
        n, out = int(self.RATE * dur), []
        for i in range(n):
            t = i / self.RATE
            env = min(1.0, t / attack) * math.exp(-decay * t)
            out.append(vol * env * sum(a * math.sin(2 * math.pi * freq * h * t)
                                       for h, a in harm))
        return out

    def _noise(self, dur, vol, lo=0.02, hi=0.22, seed=1):
        """Low-passed noise whose filter opens and closes: a soft 'whoosh'."""
        rnd, n, y, out = random.Random(seed), int(self.RATE * dur), 0.0, []
        for i in range(n):
            p = i / n
            y += (lo + (hi - lo) * math.sin(math.pi * p)) * (rnd.uniform(-1, 1) - y)
            out.append(math.sin(math.pi * p) ** 2 * y)
        peak = max(abs(x) for x in out) or 1.0
        return [x * vol / peak for x in out]

    def _mix(self, *parts):
        """parts = (start_seconds, samples) ..."""
        total = max(int(st * self.RATE) + len(x) for st, x in parts)
        buf = [0.0] * total
        for st, x in parts:
            off = int(st * self.RATE)
            for i, v in enumerate(x):
                buf[off + i] += v
        return buf

    def _write(self, name, samples):
        fade = int(0.012 * self.RATE)             # fade-out: no end click
        for i in range(min(fade, len(samples))):
            samples[-1 - i] *= i / fade
        data = array.array("h", (int(max(-1.0, min(1.0, v)) * 30000)
                                 for v in samples))
        if sys.byteorder == "big":
            data.byteswap()
        path = os.path.join(self.dir, name + ".wav")
        with wave.open(path, "wb") as w:
            w.setnchannels(1)
            w.setsampwidth(2)
            w.setframerate(self.RATE)
            w.writeframes(data.tobytes())
        self.files[name] = path

    def _build(self):
        T, W = self._tone, self._write
        C5, E5, G5, C6 = 523.25, 659.25, 783.99, 1046.5
        pure = ((1, 1.0),)
        W("click", T(1100, 0.05, 0.10, attack=0.002, decay=60, harm=pure))
        W("hint", T(880, 0.10, 0.16, decay=24, harm=pure))
        W("pop", T(784, 0.12, 0.14, decay=22, harm=((1, 1.0), (2, 0.1))))
        W("flip", self._noise(0.34, 0.14, lo=0.015, hi=0.20, seed=1))
        W("swish", self._noise(0.20, 0.08, lo=0.03, hi=0.25, seed=2))
        for k in range(6):                         # chime climbs with streak
            base = E5 * 2 ** (k / 12)
            W(f"correct{k}", self._mix((0, T(base, 0.36, 0.20)),
                                       (0.085, T(base * 1.5, 0.40, 0.20))))
        W("wrong", self._mix((0, T(329.63, 0.30, 0.20, attack=0.02, decay=9)),
                             (0.11, T(261.63, 0.38, 0.20, attack=0.02,
                                      decay=8))))
        W("finish", self._mix(*[(i * 0.12, T(f, 0.55, 0.20, decay=6))
                                for i, f in enumerate((C5, E5, G5, C6))]))
        W("done", self._mix((0, T(392.0, 0.35, 0.18, decay=8)),
                            (0.14, T(C5, 0.45, 0.18, decay=7))))

    # ---- playback --------------------------------------------------------
    def load(self):
        if not (self._loading or self.ready):
            self._loading = True
            threading.Thread(target=self._load, daemon=True).start()

    def _load(self):
        try:
            if sys.platform.startswith("win"):
                import winsound
                self._winsound = winsound
            elif sys.platform == "darwin":
                self._cmd = ["afplay"] if shutil.which("afplay") else None
            else:
                for cmd in (["paplay"], ["aplay", "-q"], ["play", "-q"]):
                    if shutil.which(cmd[0]):
                        self._cmd = cmd
                        break
            if not (self._winsound or self._cmd):
                return
            self.dir = tempfile.mkdtemp(prefix="flashcards_sfx_")
            atexit.register(shutil.rmtree, self.dir, True)
            self._build()
            self.ready = True
        except Exception:
            self.ready = False

    def play(self, name):
        path = self.files.get(name)
        if not (self.enabled and self.ready and path):
            return
        try:
            if self._winsound:
                ws = self._winsound
                ws.PlaySound(path, ws.SND_FILENAME | ws.SND_ASYNC
                             | ws.SND_NODEFAULT)
            else:
                subprocess.Popen(self._cmd + [path], stdout=subprocess.DEVNULL,
                                 stderr=subprocess.DEVNULL)
        except Exception:
            pass


SFX = Sounds()


# --------------------------------------------------------------------------
# Rounded widgets (drawn on canvases, so no sharp corners)
# --------------------------------------------------------------------------
class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, theme, kind="secondary",
                 width=120, height=42, size=11, bg_key="bg", anchor="center"):
        super().__init__(parent, width=width, height=height, bd=0,
                         highlightthickness=0, cursor="hand2")
        self.text, self.command, self.kind = text, command, kind
        self.theme, self.bg_key, self.size, self.anchor = theme, bg_key, size, anchor
        self._bw, self._bh = width, height
        self.hover = self.down = self.disabled = False
        self.indicator = None          # None = no tick-box, else True / False
        self.bind("<Enter>", lambda e: self._upd(hover=True))
        self.bind("<Leave>", lambda e: self._upd(hover=False, down=False))
        self.bind("<ButtonPress-1>", lambda e: self._upd(down=True))
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Configure>", self._resize)
        self.set_theme(theme)

    def _upd(self, **kw):
        for k, v in kw.items():
            setattr(self, k, v)
        self._draw()

    def _resize(self, e):
        self._bw, self._bh = e.width, e.height
        self._draw()

    def _release(self, e):
        was_down = self.down
        self._upd(down=False)
        if (was_down and not self.disabled and 0 <= e.x <= self._bw
                and 0 <= e.y <= self._bh and self.command):
            SFX.play("click")
            self.command()

    def set_text(self, text):
        self.text = text
        self._draw()

    def set_kind(self, kind):
        self.kind = kind
        self._draw()

    def set_indicator(self, value):
        self.indicator = value
        self._draw()

    def set_disabled(self, flag):
        self.disabled = flag
        self.configure(cursor="arrow" if flag else "hand2")
        self._draw()

    def set_theme(self, theme):
        self.theme = theme
        self.configure(bg=theme[self.bg_key])
        self._draw()

    def _draw(self):
        t = self.theme
        self.delete("all")
        if self.kind == "primary":
            fill, fg = t["accent"], t["accent_text"]
            if self.hover:
                fill = t["accent_hover"]
            if self.down:
                fill = mix(t["accent_hover"], "#000000", 0.18)
        elif self.kind in ("good", "bad"):
            fill, fg = t[self.kind], t["accent_text"]
        else:
            fill, fg = t["surface"], t["text"]
            if self.hover:
                fill = mix(t["surface"], t["text"], 0.09)
            if self.down:
                fill = mix(t["surface"], t["text"], 0.18)
        if self.disabled and self.kind in ("primary", "secondary"):
            fill, fg = t["surface"], t["muted"]
        w, h = self._bw, self._bh
        self.create_polygon(rr_points(1, 1, w - 1, h - 1, h / 2.4),
                            smooth=True, fill=fill, outline=fill)
        if self.anchor == "w":
            tx, tw = 18, max(40, w - 36)
            if self.indicator is not None:         # rounded tick-box
                bx, by = 14, (h - 20) / 2
                on = self.indicator
                self.create_polygon(
                    rr_points(bx, by, bx + 20, by + 20, 6), smooth=True,
                    fill=fg if on else t["card"],
                    outline=fg if on else t["muted"], width=2)
                if on:
                    self.create_line(bx + 5, by + 10.5, bx + 9, by + 14.5,
                                     bx + 15.5, by + 6, fill=fill, width=2.5,
                                     capstyle="round", joinstyle="round")
                tx, tw = 46, max(40, w - 64)
            size, text = fit_text(self, self.text, self.size, True, tw, h - 12)
            self.create_text(tx, h / 2, text=text, fill=fg, anchor="w",
                             width=tw, font=(FONT, size, "bold"))
        else:
            self.create_text(w / 2, h / 2, text=self.text, fill=fg,
                             font=(FONT, self.size, "bold"))


class RoundedCheck(tk.Canvas):
    """A rounded tick-box chip: click to tick / untick."""

    def __init__(self, parent, text, theme, checked=False, width=190,
                 height=38, size=11, guard=None, on_toggle=None, bg_key="bg"):
        super().__init__(parent, width=width, height=height, bd=0,
                         highlightthickness=0, cursor="hand2")
        self.text, self.checked = text, checked
        self.guard, self.on_toggle = guard, on_toggle
        self.theme, self.bg_key, self.size = theme, bg_key, size
        self._bw, self._bh = width, height
        self.hover = self.disabled = False
        self.bind("<Enter>", lambda e: self._upd(hover=True))
        self.bind("<Leave>", lambda e: self._upd(hover=False))
        self.bind("<ButtonRelease-1>", self._click)
        self.bind("<Configure>", self._resize)
        self.set_theme(theme)

    def _upd(self, **kw):
        for k, v in kw.items():
            setattr(self, k, v)
        self._draw()

    def _resize(self, e):
        self._bw, self._bh = e.width, e.height
        self._draw()

    def _click(self, e):
        if self.disabled or not (0 <= e.x <= self._bw and 0 <= e.y <= self._bh):
            return
        new = not self.checked
        if self.guard and not self.guard(self, new):
            return
        self.set(new)
        SFX.play("click")
        if self.on_toggle:
            self.on_toggle(self)

    def set(self, flag):
        self.checked = flag
        self._draw()

    def set_text(self, text):
        self.text = text
        self._draw()

    def set_disabled(self, flag):
        self.disabled = flag
        self.configure(cursor="arrow" if flag else "hand2")
        self._draw()

    def set_theme(self, theme):
        self.theme = theme
        self.configure(bg=theme[self.bg_key])
        self._draw()

    def _draw(self):
        t, w, h = self.theme, self._bw, self._bh
        self.delete("all")
        on = self.checked and not self.disabled
        if self.disabled:
            fill = edge = t["surface"]
            fg = t["muted"]
        elif self.checked:
            fill, edge, fg = mix(t["surface"], t["accent"], 0.22), t["accent"], t["text"]
        else:
            fill = mix(t["surface"], t["text"], 0.07) if self.hover else t["surface"]
            edge, fg = fill, t["text"]
        self.create_polygon(rr_points(1, 1, w - 1, h - 1, h / 2.4),
                            smooth=True, fill=fill, outline=edge, width=2)
        bx, by, bs = 14, (h - 18) / 2, 18
        self.create_polygon(
            rr_points(bx, by, bx + bs, by + bs, 6), smooth=True,
            fill=t["accent"] if on else t["card"],
            outline=t["accent"] if on else t["muted"], width=2)
        if on:
            self.create_line(bx + 4.5, by + 9.5, bx + 8, by + 13,
                             bx + 13.5, by + 5, fill=t["accent_text"],
                             width=2.5, capstyle="round", joinstyle="round")
        room = max(20, w - bx - bs - 26)          # shrink the font to fit
        size, text = fit_text(self, self.text, self.size, on, room, h - 8,
                              minimum=8)
        self.create_text(bx + bs + 12, h / 2, text=text, fill=fg, anchor="w",
                         width=room,
                         font=(FONT, size, "bold" if on else "normal"))


class RoundedSelect(RoundedButton):
    """A rounded drop-down: click to pick from a menu of options."""

    def __init__(self, parent, theme, variable, get_options, width=300,
                 on_change=None, bg_key="bg"):
        self.var, self.get_options, self.on_change = variable, get_options, on_change
        super().__init__(parent, "", self._open, theme, width=width,
                         height=44, size=11, bg_key=bg_key, anchor="w")
        self.refresh()

    def refresh(self):
        self.set_text(f"{self.var.get()}      \u25be")

    def _open(self):
        def pick(o):
            self.var.set(o)
            self.refresh()
            if self.on_change:
                self.on_change(o)
        post_menu(self, self.theme, self.get_options(), pick)


class RoundedEntry(tk.Canvas):
    def __init__(self, parent, theme, textvariable=None, width=320, height=44,
                 size=14, justify="left", bg_key="bg", digits_only=False):
        super().__init__(parent, width=width, height=height, bd=0,
                         highlightthickness=0)
        self.theme, self.bg_key, self.focused = theme, bg_key, False
        self._bw, self._bh = width, height
        self.entry = tk.Entry(self, textvariable=textvariable, bd=0,
                              relief="flat", highlightthickness=0,
                              font=(FONT, size), justify=justify)
        self._win = self.create_window(width / 2, height / 2,
                                       window=self.entry, width=width - 30,
                                       height=height - 16)
        self.entry.bind("<FocusIn>", lambda e: self._focus(True))
        self.entry.bind("<FocusOut>", lambda e: self._focus(False))
        self.bind("<Button-1>", lambda e: self.entry.focus_set())
        self.bind("<Configure>", self._resize)
        if digits_only:                        # numbers only, up to 3 digits
            self.entry.configure(
                validate="key",
                validatecommand=(self.register(
                    lambda p: p == "" or (p.isdigit() and len(p) <= 3)), "%P"))
        self.set_theme(theme)

    def on_blur(self, fn):
        self.entry.bind("<FocusOut>", lambda e: fn(), add="+")

    def _focus(self, flag):
        self.focused = flag
        self._draw()

    def _resize(self, e):
        self._bw, self._bh = e.width, e.height
        self.coords(self._win, e.width / 2, e.height / 2)
        self.itemconfigure(self._win, width=max(10, e.width - 30),
                           height=max(10, e.height - 16))
        self._draw()

    def get(self):
        return self.entry.get()

    def focus_field(self):
        self.entry.focus_set()

    def enable(self, flag):
        self.entry.configure(state="normal" if flag else "disabled")

    def set_theme(self, theme):
        self.theme = theme
        self.configure(bg=theme[self.bg_key])
        self.entry.configure(
            bg=theme["entry_bg"], fg=theme["text"],
            insertbackground=theme["text"],
            disabledbackground=theme["entry_bg"],
            disabledforeground=theme["muted"],
            selectbackground=theme["accent"],
            selectforeground=theme["accent_text"])
        self._draw()

    def _draw(self):
        t = self.theme
        self.delete("bg")
        border = t["accent"] if self.focused else t["entry_border"]
        self.create_polygon(rr_points(2, 2, self._bw - 2, self._bh - 2,
                                      self._bh / 2.6),
                            smooth=True, fill=t["entry_bg"], outline=border,
                            width=2, tags="bg")
        self.tag_lower("bg")


class ProgressBar(tk.Canvas):
    def __init__(self, parent, theme, height=10, bg_key="bg"):
        super().__init__(parent, height=height, bd=0, highlightthickness=0)
        self.theme, self.bg_key, self.value, self._cancel = theme, bg_key, 0.0, None
        self.bind("<Configure>", lambda e: self._draw())
        self.set_theme(theme)

    def set(self, v, animated=True):
        if self._cancel:
            self._cancel()
        a, b = self.value, v
        if not animated:
            self.value = b
            self._draw()
            return

        def step(p):
            self.value = a + (b - a) * p
            self._draw()
        self._cancel = animate(self, 380, step, ease=ease_out)

    def set_theme(self, theme):
        self.theme = theme
        self.configure(bg=theme[self.bg_key])
        self._draw()

    def _draw(self):
        t = self.theme
        w, h = self.winfo_width(), self.winfo_height()
        if w < 4:
            return
        self.delete("all")
        self.create_polygon(rr_points(0, 0, w, h, h / 2), smooth=True,
                            fill=t["surface"], outline=t["surface"])
        if self.value > 0.001:
            self.create_polygon(rr_points(0, 0, max(h, w * self.value), h,
                                          h / 2), smooth=True,
                                fill=t["accent"], outline=t["accent"])


# --------------------------------------------------------------------------
# The flash card itself: flips, slides, shakes and pulses
# --------------------------------------------------------------------------
class CardView(tk.Canvas):
    def __init__(self, parent, theme):
        super().__init__(parent, bd=0, highlightthickness=0, height=200)
        self.theme = theme
        self.face = dict(tag="", category="", text="", sub="", hint="",
                         question="", extra="", fill="card",
                         outline="card_border")
        self.sx, self.offset, self.pop = 1.0, 0.0, 1.0
        self._cancel = None
        self._lay_key, self._lay = None, None
        self.bind("<Configure>", lambda e: self.draw())
        self.set_theme(theme)

    # ---- drawing ---------------------------------------------------------
    @staticmethod
    def _font_size(text):
        n = len(text)
        return (30 if n <= 24 else 26 if n <= 48 else
                22 if n <= 90 else 18 if n <= 160 else 15)

    def draw(self):
        W, H = self.winfo_width(), self.winfo_height()
        if W < 20 or H < 20:
            return
        t, f = self.theme, self.face
        self.delete("all")
        cw, ch = min(W - 40, 640), max(120, H - 36)
        cx, cy = W / 2 + self.offset, H / 2 - 4
        hw = max(1.0, cw / 2 * self.sx * self.pop)
        hh = ch / 2 * (0.93 + 0.07 * self.sx) * self.pop
        self.create_polygon(rr_points(cx - hw, cy - hh + 8, cx + hw,
                                      cy + hh + 8, 30),
                            smooth=True, fill=t["shadow"], outline="")
        self.create_polygon(rr_points(cx - hw, cy - hh, cx + hw, cy + hh, 30),
                            smooth=True, fill=t[f["fill"]],
                            outline=t[f["outline"]], width=3)
        if self.sx < 0.8:          # edge-on: hide text while turning
            return
        label = f["tag"] + (f"   \u00b7   {f['category']}" if f["category"] else "")
        self.create_text(cx, cy - hh + 30, text=label, fill=t["muted"],
                         font=(FONT, 10, "bold"))
        colors = {"q": t["muted"], "m": t["text"], "e": t["accent"],
                  "s": t["muted"]}
        for name, text, size, bold, y in self._layout(cw, ch):
            self.create_text(cx, cy + y, text=text, fill=colors[name],
                             width=cw - 90, anchor="n", justify="center",
                             font=(FONT, size, "bold" if bold else "normal"))
        if f["hint"]:
            self.create_text(cx, cy + hh - 32, text=f["hint"], fill=t["warn"],
                             font=("Consolas", 15, "bold"))

    def _layout(self, cw, ch):
        """Stack question / answer / proposed answer / note without overlap,
        shrinking the text until everything fits inside the card."""
        f = self.face
        key = (cw, ch, f["text"], f["question"], f["extra"], f["sub"],
               bool(f["hint"]), FONT)
        if key == self._lay_key:
            return self._lay
        wrap, gap = cw - 90, 12
        top = -ch / 2 + 50
        band = max(60, (ch / 2 - (62 if f["hint"] else 26)) - top)
        main = self._font_size(f["text"])
        if f["question"]:
            main = min(main, 26)
        elif f["extra"]:
            main = min(main, 24)
        if "\n" in f["text"]:
            main = min(main, 18)
        parts = []                          # [name, text, size, min, bold]
        if f["question"]:
            parts.append(["q", f["question"], 13, 9, False])
        parts.append(["m", f["text"], main, 10, True])
        if f["extra"]:
            parts.append(["e", f["extra"], 20, 11, True])
        if f["sub"]:
            parts.append(["s", f["sub"], 12, 9, False])

        def heights():
            return [wrap_height(self, p[1], p[2], p[4], wrap) for p in parts]
        hs = heights()
        while sum(hs) + gap * (len(parts) - 1) > band:
            shrinkable = [i for i, p in enumerate(parts) if p[2] > p[3]]
            if not shrinkable:
                break
            i = max(shrinkable, key=lambda i: hs[i])
            parts[i][2] -= 1
            hs = heights()
        y = top + max(0, (band - sum(hs) - gap * (len(parts) - 1)) / 2)
        out = []
        for p, h in zip(parts, hs):
            out.append((p[0], p[1], p[2], p[4], y))
            y += h + gap
        self._lay_key, self._lay = key, out
        return out

    def set_theme(self, theme):
        self.theme = theme
        self.configure(bg=theme["bg"])
        self.draw()

    def set_face(self, **face):
        self.face.update(face)
        self.draw()

    # ---- animations ------------------------------------------------------
    def _run(self, ms, step, done=None, ease=ease_in_out):
        if self._cancel:
            self._cancel()
        self._cancel = animate(self, ms, step, done, ease)

    def flip(self, new_face, done=None):
        """Turn the card over (squeeze to an edge, swap face, widen again)."""
        swapped = {"v": False}

        def step(p):
            self.sx = abs(math.cos(math.pi * p))
            if p >= 0.5 and not swapped["v"]:
                swapped["v"] = True
                self.face.update(new_face)
            self.draw()

        def finish():
            self.sx = 1.0
            self.draw()
            if done:
                done()
        self._run(520, step, finish, ease_in_out)

    def slide(self, start, end, ms, done=None, ease=ease_out):
        def step(p):
            self.offset = start + (end - start) * p
            self.draw()
        self._run(ms, step, done, ease)

    def shake(self):
        def step(p):
            self.offset = math.sin(p * math.pi * 7) * 16 * (1 - p)
            self.draw()

        def finish():
            self.offset = 0
            self.draw()
        self._run(420, step, finish, lambda p: p)

    def pulse(self):
        def step(p):
            self.pop = 1 + 0.05 * math.sin(math.pi * p)
            self.draw()

        def finish():
            self.pop = 1.0
            self.draw()
        self._run(380, step, finish, lambda p: p)


# --------------------------------------------------------------------------
# The application
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, store):
        global FONT
        super().__init__()
        self.store = store
        load_bundled_fonts()
        FONT = pick_font(tkfont.families(self))
        self.title("Flash Card Game")
        self._set_window_icon()
        # fit the window to the screen (small laptops hide the bottom edge)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = min(840, sw - 40), max(560, min(720, sh - 130))
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}"
                      f"+{max(0, (sh - h) // 2 - 20)}")
        self.minsize(min(720, w), min(580, h))

        self.theme_name = ("dark" if self._get_setting("theme", "light") == "dark"
                           else "light")
        self.theme = dict(THEMES[self.theme_name])
        self._themed, self._theme_cancel = [], None
        self.configure(bg=self.theme["bg"])

        # game state
        self.mode, self.busy = "menu", False
        self.deck, self.idx = [], 0
        self.correct = self.streak = self.best = self.played = 0
        self.player, self.cat_label = "Player", "All"
        self.round_start, self.total_time = None, 0.0
        self.paused, self.pause_started, self.paused_total = False, 0.0, 0.0
        self.answered = self.hinted = self.last_wrong = False
        self.prev_streak = 0
        self.pending = None
        self._timer_job = None
        self.fb_key = "muted"
        self.pool_cards, self.card_modes = [], []
        self.cur_mode, self.mode_label = "id", MODE_NAMES["id"]
        self.choices, self.correct_idx, self.picked = [], 0, None
        self.right_idx, self.selected, self.tf_proposed = set(), set(), ""
        self.removed = set()
        self.review, self.missed, self.msgs = False, [], {}
        self.res_badge_key = "good"
        self.sel_cats, self.sel_modes = set(), set()
        self.count_custom = False
        self.cat_checks, self.mode_checks, self._cat_list = {}, {}, []

        self._load_prefs()
        SFX.enabled = self._get_setting("sound", "on") != "off"
        SFX.load()
        self._build_header()
        content = self.mkframe(self)
        content.pack(fill="both", expand=True)
        content.rowconfigure(0, weight=1)
        content.columnconfigure(0, weight=1)
        self.frames = {}
        for name, build in (("menu", self._build_menu),
                            ("game", self._build_game),
                            ("results", self._build_results)):
            f = self.mkframe(content)
            f.grid(row=0, column=0, sticky="nsew")
            build(f)
            self.frames[name] = f

        self.bind("<Return>", self.on_enter)
        self.bind("<Key>", self.on_key)
        self.bind("<Escape>", lambda e: self.toggle_pause())
        self.protocol("WM_DELETE_WINDOW", self.quit_app)
        self.show("menu")

    # ------------------------------------------------------------------
    # theming
    # ------------------------------------------------------------------
    def _set_window_icon(self):
        """Title-bar / taskbar icon: icon.ico bundled in the exe or beside it."""
        if not sys.platform.startswith("win"):
            return
        for base in (Path(getattr(sys, "_MEIPASS", app_dir())), app_dir()):
            ico = base / "flashcards.ico"
            if ico.exists():
                try:
                    self.iconbitmap(default=str(ico))
                except tk.TclError:
                    pass
                return

    def reg(self, fn):
        self._themed.append(fn)
        fn(self.theme)

    def cfg(self, widget, **opts):
        """Keep widget options bound to theme keys, e.g. cfg(w, bg='bg')."""
        self.reg(lambda t, w=widget, o=opts:
                 w.configure(**{k: t[v] for k, v in o.items()}))
        return widget

    def reg_w(self, widget):
        self.reg(widget.set_theme)
        return widget

    def mkframe(self, parent, bg="bg"):
        return self.cfg(tk.Frame(parent, bd=0, highlightthickness=0), bg=bg)

    def lbl(self, parent, text="", size=11, bold=False, fg="text", bg="bg"):
        w = tk.Label(parent, text=text, bd=0,
                     font=(FONT, size, "bold" if bold else "normal"))
        return self.cfg(w, fg=fg, bg=bg)

    def apply_theme(self, theme):
        self.theme = theme
        self.configure(bg=theme["bg"])
        for fn in list(self._themed):
            try:
                fn(theme)
            except tk.TclError:
                pass

    def toggle_theme(self):
        target = "light" if self.theme_name == "dark" else "dark"
        start, end = dict(self.theme), THEMES[target]
        if self._theme_cancel:
            self._theme_cancel()
        self.theme_name = target
        self._update_toggle_text()
        self._theme_cancel = animate(
            self, 280, lambda p: self.apply_theme(mix_theme(start, end, p)),
            ease=ease_in_out)
        self.after(400, lambda: self._set_setting("theme", target))

    def toggle_sound(self):
        SFX.enabled = not SFX.enabled
        self._update_sound_text()
        SFX.play("pop")
        self.after(400, lambda: self._set_setting(
            "sound", "on" if SFX.enabled else "off"))

    def _update_sound_text(self):
        self.sound_btn.set_text("Sound on" if SFX.enabled else "Sound off")

    def _update_toggle_text(self):
        self.theme_btn.set_text("Light mode" if self.theme_name == "dark"
                                else "Dark mode")

    # ---- settings kept in the workbook ("Settings" sheet) ------------------
    def _get_setting(self, key, default):
        wb = self.store.wb
        if "Settings" in wb.sheetnames:
            for row in wb["Settings"].iter_rows(min_row=2, values_only=True):
                if row and row[0] == key and row[1] is not None:
                    return row[1]
        return default

    def _set_setting(self, key, value, save=True):
        wb = self.store.wb
        ws = (wb["Settings"] if "Settings" in wb.sheetnames
              else wb.create_sheet("Settings"))
        if ws["A1"].value is None:
            ws["A1"], ws["B1"] = "Setting", "Value"
        for row in ws.iter_rows(min_row=2):
            if row[0].value == key:
                row[1].value = value
                break
        else:
            ws.append([key, value])
        if save:
            self._flush()

    def _flush(self):
        try:
            self.store.wb.save(self.store.path)
        except PermissionError:
            pass            # file open in Excel; saved with the next game

    def _load_prefs(self):
        """Menu choices from last time (nothing is ticked on first start)."""
        g = self._get_setting
        split = lambda v: {x for x in str(v or "").split("|") if x}
        self.sel_cats = split(g("sel_cats", ""))
        self.sel_modes = split(g("sel_modes", "")) & set(MODE_KEYS)
        count = str(g("count", "max"))
        self.count_custom = count.isdigit()
        self.pref_count = count if count.isdigit() else "10"
        self.pref_name = str(g("player", "Player"))

    def _save_prefs(self, flush=True):
        """Remember the menu choices in the workbook's Settings sheet."""
        put = lambda k, v: self._set_setting(k, v, save=False)
        put("player", self.name_var.get().strip() or "Player")
        put("sel_cats", "|".join(sorted(self.sel_cats)))
        put("sel_modes", "|".join(sorted(self.sel_modes)))
        put("count", self.count_var.get() if self.count_custom else "max")
        if flush:
            self._flush()

    def save(self):
        """Write the workbook; offer a retry if it's open in Excel."""
        while True:
            try:
                self.store.wb.save(self.store.path)
                return True
            except PermissionError:
                if not messagebox.askretrycancel(
                        "Can't save",
                        f"'{self.store.path.name}' is open in Excel.\n"
                        "Close it, then click Retry."):
                    return False

    def show(self, name):
        self.mode = name
        self.frames[name].tkraise()
        if name == "menu":
            self.refresh_menu()

    # ------------------------------------------------------------------
    # header (theme switch)
    # ------------------------------------------------------------------
    def _build_header(self):
        bar = self.mkframe(self)
        bar.pack(fill="x")
        self.theme_btn = self.reg_w(RoundedButton(
            bar, "", self.toggle_theme, self.theme, width=130, height=36,
            size=10))
        self.theme_btn.pack(side="right", padx=(8, 22), pady=(14, 0))
        self.sound_btn = self.reg_w(RoundedButton(
            bar, "", self.toggle_sound, self.theme, width=110, height=36,
            size=10))
        self.sound_btn.pack(side="right", pady=(14, 0))
        self._update_toggle_text()
        self._update_sound_text()

    # ------------------------------------------------------------------
    # menu screen
    # ------------------------------------------------------------------
    def _build_menu(self, f):
        f.columnconfigure(0, weight=1)
        self.lbl(f, "Flash Cards", 28, True, fg="accent").grid(row=0,
                                                                pady=(4, 0))
        self.deck_info = self.lbl(f, "", 11, fg="muted")
        self.deck_info.grid(row=1, pady=(0, 6))

        top = self.mkframe(f)
        top.grid(row=2, pady=(0, 2))
        self.name_var = tk.StringVar(value=self.pref_name)
        self.count_var = tk.StringVar(value=self.pref_count)
        self.lbl(top, "Name", 11, fg="muted").grid(row=0, column=0,
                                                   padx=(0, 8))
        self.reg_w(RoundedEntry(top, self.theme, self.name_var, width=220,
                                height=42, size=12)).grid(row=0, column=1)
        self.lbl(top, "Cards", 11, fg="muted").grid(row=0, column=2,
                                                    padx=(30, 8))
        step = self.mkframe(top)
        step.grid(row=0, column=3)
        self.reg_w(RoundedButton(step, "\u2212", lambda: self._step(-1),
                                 self.theme, width=42, height=42,
                                 size=15)).pack(side="left")
        self.count_entry = self.reg_w(RoundedEntry(
            step, self.theme, self.count_var, width=70, height=42, size=14,
            justify="center", digits_only=True))
        self.count_entry.pack(side="left", padx=6)
        self.count_entry.on_blur(self._on_count_edited)
        self.reg_w(RoundedButton(step, "+", lambda: self._step(1),
                                 self.theme, width=42, height=42,
                                 size=15)).pack(side="left")
        self.pool_lbl = self.lbl(step, "", 11, fg="muted")
        self.pool_lbl.pack(side="left", padx=(8, 0))

        head = self.mkframe(f)
        head.grid(row=3, sticky="ew", padx=110, pady=(8, 0))
        head.columnconfigure(0, weight=1)
        self.lbl(head, "CATEGORIES", 10, True, fg="muted").grid(
            row=0, column=0, sticky="w")
        self.cat_all = self.reg_w(RoundedCheck(
            head, "Select all", self.theme, width=124, height=28, size=10,
            on_toggle=self._toggle_all_cats))
        self.cat_all.grid(row=0, column=1)
        self.cat_frame = self.mkframe(f)
        self.cat_frame.grid(row=4)

        head2 = self.mkframe(f)
        head2.grid(row=5, sticky="ew", padx=110, pady=(6, 0))
        head2.columnconfigure(0, weight=1)
        self.lbl(head2, "MODES", 10, True, fg="muted").grid(
            row=0, column=0, sticky="w")
        self.mode_all = self.reg_w(RoundedCheck(
            head2, "Select all", self.theme, width=124, height=28, size=10,
            on_toggle=self._toggle_all_modes))
        self.mode_all.grid(row=0, column=1)
        self.mode_frame = self.mkframe(f)
        self.mode_frame.grid(row=6)
        self.mode_checks = {}
        for i, key in enumerate(MODE_KEYS):
            chip = self.reg_w(RoundedCheck(
                self.mode_frame, MODE_NAMES[key], self.theme,
                checked=(key in self.sel_modes), width=300, height=34,
                on_toggle=self._on_sel_change))
            chip.grid(row=i // 2, column=i % 2, padx=4, pady=3)
            self.mode_checks[key] = chip

        self.reg_w(RoundedButton(f, "Start Game", self.start_game, self.theme,
                                 kind="primary", width=280, height=50,
                                 size=15)).grid(row=7, pady=(10, 8))
        row = self.mkframe(f)
        row.grid(row=8)
        for i, (text, cmd) in enumerate((
                ("Leaderboard", self.show_leaderboard),
                ("Hardest cards", self.show_hardest),
                ("Add card", self.add_card_dialog),
                ("Quit", self.quit_app))):
            self.reg_w(RoundedButton(row, text, cmd, self.theme, width=134,
                                     height=38)).grid(row=0, column=i, padx=5)

    # ---- menu helpers ----------------------------------------------------
    def _categories(self):
        return sorted({c["category"] for c in self.store.get_cards()}
                      | {c["category"] for c in self.store.get_cata()})

    def _rebuild_categories(self):
        cats = self._categories()
        for w in self.cat_frame.winfo_children():
            w.destroy()
        self.cat_checks = {}
        keep = self.sel_cats & set(cats)          # nothing is ticked by default
        cols = 3 if len(cats) <= 12 else 4
        width = 196 if cols == 3 else 150
        for i, cat in enumerate(cats):
            chip = self.reg_w(RoundedCheck(
                self.cat_frame, cat, self.theme, checked=(cat in keep),
                width=width, height=34, on_toggle=self._on_sel_change))
            chip.grid(row=i // cols, column=i % cols, padx=4, pady=3)
            self.cat_checks[cat] = chip
        self._cat_list, self.sel_cats = cats, set(keep)

    def _toggle_all_cats(self, chip):
        for c in self.cat_checks.values():
            c.set(chip.checked)
        self._on_sel_change()

    def _toggle_all_modes(self, chip):
        for c in self.mode_checks.values():
            if not c.disabled:
                c.set(chip.checked)
        self._on_sel_change()

    def _on_sel_change(self, _chip=None):
        self.sel_cats = {c for c, ch in self.cat_checks.items() if ch.checked}
        self.sel_modes = {k for k, ch in self.mode_checks.items()
                          if ch.checked and not ch.disabled}
        cats = list(self.cat_checks.values())
        modes = [c for c in self.mode_checks.values() if not c.disabled]
        self.cat_all.set(bool(cats) and all(c.checked for c in cats))
        self.mode_all.set(bool(modes) and all(c.checked for c in modes))
        self._clamp_count()

    def _pool(self):
        """Every question that matches the ticked categories and modes."""
        items = []
        if self.sel_modes & {"id", "mc", "tf"}:
            items += [c for c in self.store.get_cards()
                      if c["category"] in self.sel_cats]
        if "cata" in self.sel_modes:
            items += [c for c in self.store.get_cata()
                      if c["category"] in self.sel_cats]
        return items

    def _clamp_count(self):
        """Card count: the maximum available, unless the player chose one."""
        total = len(self._pool())
        if total == 0:
            n = 0
        elif not self.count_custom:
            n = total
        else:
            try:
                n = int(self.count_var.get())
            except ValueError:
                n = total
            n = max(1, min(n, total))
        self.count_var.set(str(n))
        self.pool_lbl.config(text=f"of {total}")

    def _on_count_edited(self):
        total = len(self._pool())
        try:
            n = int(self.count_var.get())
        except ValueError:
            n = total
        n = max(1, min(n, total)) if total else 0
        self.count_custom = (n != total)
        self.count_var.set(str(n))
        self._clamp_count()

    def _step(self, d):
        total = len(self._pool())
        try:
            n = int(self.count_var.get())
        except ValueError:
            n = total
        n = max(1, min(n + d, total)) if total else 0
        self.count_custom = (n != total)
        self.count_var.set(str(n))
        self._clamp_count()

    def refresh_menu(self):
        cards, cata = self.store.get_cards(), self.store.get_cata()
        if self._categories() != self._cat_list:
            self._rebuild_categories()
        chip = self.mode_checks["cata"]
        chip.set_disabled(not cata)
        chip.set_text("Check all that apply" if cata
                      else "Check all that apply  (no questions yet)")
        if not cata and chip.checked:
            chip.set(False)
        self.deck_info.config(
            text=f"{len(cards)} cards   \u00b7   {len(cata)} check-all "
                 f"questions   \u00b7   {self.store.path.name}")
        self._on_sel_change()

    # ------------------------------------------------------------------
    # game screen
    # ------------------------------------------------------------------
    def _build_game(self, f):
        f.columnconfigure(0, weight=1)
        f.rowconfigure(2, weight=1)

        top = self.mkframe(f)
        top.grid(row=0, sticky="ew", padx=30, pady=(4, 6))
        top.columnconfigure(1, weight=1)
        self.lbl_progress = self.lbl(top, "", 12, True)
        self.lbl_progress.grid(row=0, column=0, sticky="w")
        self.lbl_timer = self.lbl(top, "Time 0:00", 12, fg="muted")
        self.lbl_timer.grid(row=0, column=1)
        self.lbl_score = self.lbl(top, "", 15, True, fg="accent")
        self.lbl_score.grid(row=0, column=2, padx=(0, 14))
        self.reg_w(RoundedButton(top, "Pause", self.toggle_pause, self.theme,
                                 width=80, height=34, size=10)).grid(
            row=0, column=3, padx=(0, 6))
        self.reg_w(RoundedButton(top, "End round", self.end_round, self.theme,
                                 width=100, height=34, size=10)).grid(
            row=0, column=4)

        self.progress = self.reg_w(ProgressBar(f, self.theme))
        self.progress.grid(row=1, sticky="ew", padx=30, pady=(0, 6))

        self.card = self.reg_w(CardView(f, self.theme))
        self.card.grid(row=2, sticky="nsew", padx=20)

        self.input_area = self.mkframe(f)
        self.input_area.grid(row=3, sticky="ew", padx=60, pady=(8, 2))
        self.input_area.columnconfigure(0, weight=1)
        self.answer_var = tk.StringVar()
        self.entry = self.reg_w(RoundedEntry(
            self.input_area, self.theme, self.answer_var, height=50, size=18,
            justify="center"))
        self.entry.grid(row=0, column=0, sticky="ew", padx=50)
        self.mc_frame = self.mkframe(self.input_area)
        self.mc_frame.grid(row=0, column=0, sticky="ew")
        self.mc_frame.columnconfigure((0, 1), weight=1, uniform="mc")
        self.mc_btns = []
        for i in range(MAX_CATA_OPTIONS):
            b = self.reg_w(RoundedButton(
                self.mc_frame, "", lambda i=i: self.choose(i), self.theme,
                width=260, height=46, size=12, anchor="w"))
            b.grid(row=i // 2, column=i % 2, sticky="ew", padx=5, pady=3)
            self.mc_btns.append(b)
        self.mc_frame.grid_remove()

        self.lbl_feedback = tk.Label(f, bd=0, font=(FONT, 13, "bold"),
                                     justify="center", wraplength=720)
        self.reg(lambda t: self.lbl_feedback.configure(
            bg=t["bg"], fg=t[self.fb_key]))
        self.lbl_feedback.grid(row=4, sticky="ew", padx=34, pady=(2, 2))

        btns = self.mkframe(f)
        btns.grid(row=5, pady=(2, 10))
        mk = lambda text, cmd, kind="secondary", w=140: self.reg_w(
            RoundedButton(btns, text, cmd, self.theme, kind=kind, width=w,
                          height=44, size=12))
        self.btn_submit = mk("Submit", self.submit, "primary", 150)
        self.btn_hint = mk("Hint", self.hint, w=110)
        self.btn_skip = mk("Show answer", self.skip, w=150)
        self.btn_override = mk("I was right", self.override, w=150)
        self.btn_next = mk("Next card", self.next_card, "primary", 170)
        self._btns = (self.btn_submit, self.btn_hint, self.btn_skip,
                      self.btn_override, self.btn_next)

        # pause cover: hides the question while the clock is stopped
        self.pause_overlay = self.mkframe(f, bg="surface")
        box = self.mkframe(self.pause_overlay, bg="surface")
        box.place(relx=0.5, rely=0.45, anchor="center")
        self.lbl(box, "Paused", 34, True, fg="accent", bg="surface").pack()
        self.lbl(box, "Your time is stopped and the question is hidden.", 12,
                 fg="muted", bg="surface").pack(pady=(4, 18))
        self.reg_w(RoundedButton(box, "Resume", self.toggle_pause, self.theme,
                                 kind="primary", width=200, height=52,
                                 size=14, bg_key="surface")).pack()

    def _set_buttons(self, mode):
        for b in self._btns:
            b.grid_remove()
        if mode == "ask":
            col = 0
            if self.cur_mode in ("id", "cata"):
                self.btn_submit.grid(row=0, column=0, padx=5)
                col = 1
            self.btn_hint.grid(row=0, column=col, padx=5)
            self.btn_skip.grid(row=0, column=col + 1, padx=5)
            self.btn_hint.set_disabled(
                self.hinted or self.cur_mode == "tf"
                or (self.cur_mode == "mc" and len(self.choices) <= 2))
        elif mode in ("review_ok", "review_wrong"):
            if mode == "review_wrong":
                self.btn_override.grid(row=0, column=0, padx=5)
            self.btn_next.grid(row=0, column=1, padx=5)

    def _feedback(self, text, key):
        self.fb_key = key
        self.lbl_feedback.configure(text=text, fg=self.theme[key])

    def _update_score(self):
        self.lbl_score.config(
            text=f"Correct  {self.correct} / {len(self.deck)}")

    # ---- the round clock (one timer for the whole round) --------------------
    @staticmethod
    def _fmt_time(sec):
        sec = int(sec)
        return f"{sec // 60}:{sec % 60:02d}"

    def _elapsed(self):
        if self.round_start is None:
            return 0.0
        now = self.pause_started if self.paused else time.perf_counter()
        return max(0.0, now - self.round_start - self.paused_total)

    def _tick(self):
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None
        if self.mode != "game":
            return
        self.lbl_timer.config(text=f"Time {self._fmt_time(self._elapsed())}")
        self._timer_job = self.after(250, self._tick)

    def toggle_pause(self):
        if self.mode != "game":
            return
        if self.paused:
            self._resume()
        else:
            self._pause()

    def _pause(self, cover=True):
        if self.paused:
            return
        self.paused, self.pause_started = True, time.perf_counter()
        if cover:
            self.pause_overlay.place(x=0, y=0, relwidth=1, relheight=1)
            self.pause_overlay.lift()
            SFX.play("pop")

    def _resume(self, sound=True):
        if not self.paused:
            return
        if self.round_start is not None:
            self.paused_total += time.perf_counter() - self.pause_started
        self.paused = False
        self.pause_overlay.place_forget()
        if sound:
            SFX.play("pop")
        self._tick()
        if self.cur_mode == "id" and not self.answered:
            self.entry.focus_field()
        else:
            self.focus_set()

    # ---- starting / loading cards -------------------------------------------
    def start_game(self):
        self._on_count_edited()
        if not self.sel_cats or not self.sel_modes:
            messagebox.showinfo(
                "Nothing selected",
                "Tick at least one category and one mode to start.")
            return
        items = self._pool()
        if not items:
            messagebox.showinfo(
                "No questions",
                "There are no questions for the ticked categories and modes.\n"
                "Tick other options, or add questions in the Excel file.")
            return
        try:
            n = int(self.count_var.get())
        except ValueError:
            n = len(items)
        n = max(1, min(n, len(items)))
        self.player = self.name_var.get().strip() or "Player"
        self._save_prefs()

        cats = sorted(self.sel_cats)
        self.cat_label = ("All" if len(cats) == len(self.cat_checks)
                          else ", ".join(cats[:3]) + ("\u2026" if len(cats) > 3
                                                     else ""))
        keys = [k for k in MODE_KEYS if k in self.sel_modes]
        self.mode_label = (MODE_NAMES[keys[0]] if len(keys) == 1 else
                           "Mixed: " + ", ".join(MODE_SHORT[k] for k in keys))
        card_modes = [k for k in keys if k != "cata"]

        deck = pick_questions(items, n)
        modes = ["cata" if it["type"] == "cata"
                 else random.choice(card_modes) for it in deck]
        self._begin_round(deck, modes, review=False)

    def start_review(self):
        """Retry only the questions that were answered wrong last round."""
        if not self.missed:
            return
        pairs = self.missed[:]
        random.shuffle(pairs)
        self._begin_round([p[0] for p in pairs], [p[1] for p in pairs],
                          review=True)

    def _begin_round(self, deck, modes, review):
        self.deck, self.card_modes, self.review = deck, modes, review
        self.pool_cards = self.store.get_cards()      # source of auto-choices
        self.msgs = self.store.get_messages()
        self.missed = []
        self.idx = 0
        self.correct = self.streak = self.best = self.played = 0
        self.pending = None
        self.round_start, self.paused_total, self.total_time = None, 0.0, 0.0
        self.paused = False
        self.pause_overlay.place_forget()
        self.lbl_timer.config(text="Time 0:00")
        self.progress.set(0, animated=False)
        self.show("game")
        self._load_card(slide=True)

    # ---- building the options for each mode ---------------------------------
    def _wrong_pool(self, item):
        """Wrong answers for a card: the ones typed in Excel, else borrowed."""
        correct = item["answer"].strip()
        seen, admin = {correct.lower()}, []
        for w in item.get("wrong", []):
            if w.lower() not in seen:
                seen.add(w.lower())
                admin.append(w)
        if admin:
            return admin
        same, other = [], []
        for c in self.pool_cards:
            a = c["answer"].strip()
            if a.lower() in seen:
                continue
            seen.add(a.lower())
            (same if c["category"] == item["category"] else other).append(a)
        random.shuffle(same)
        random.shuffle(other)
        return same + other

    def _cata_options(self, item):
        right, wrong = item["right"][:], item["wrong"][:]
        random.shuffle(right)
        random.shuffle(wrong)
        n_r = min(len(right), 3)
        pr, pw = right[:n_r], wrong[:MAX_CATA_OPTIONS - n_r]
        i = n_r
        while len(pr) + len(pw) < MAX_CATA_OPTIONS and i < len(right):
            pr.append(right[i])
            i += 1
        opts = [(x, True) for x in pr] + [(x, False) for x in pw]
        random.shuffle(opts)
        return opts

    def _prepare(self, item, mode):
        """Set up choices for this card; returns the mode actually used."""
        self.choices, self.right_idx, self.tf_proposed = [], set(), ""
        if mode == "cata":
            opts = self._cata_options(item)
            self.choices = [t for t, _ in opts]
            self.right_idx = {i for i, (_, r) in enumerate(opts) if r}
            return "cata"
        correct = item["answer"].strip()
        if mode in ("mc", "tf"):
            pool = self._wrong_pool(item)
            if mode == "mc" and pool:
                opts = pool[:3] + [correct]
                random.shuffle(opts)
                self.choices, self.correct_idx = opts, opts.index(correct)
                return "mc"
            if mode == "tf" and pool:
                is_true = random.random() < 0.5
                self.tf_proposed = correct if is_true else random.choice(pool)
                self.choices = ["True", "False"]
                self.correct_idx = 0 if is_true else 1
                return "tf"
        return "id"                 # fall back to typing

    def _show_mode_ui(self):
        if self.cur_mode == "id":
            self.mc_frame.grid_remove()
            self.entry.grid()
            return
        self.entry.grid_remove()
        self.mc_frame.grid()
        for b in self.mc_btns:               # roomier buttons when few options
            b.configure(height=56 if len(self.choices) <= 4 else 46)
        for i, b in enumerate(self.mc_btns):
            if i < len(self.choices):
                b.set_kind("secondary")
                b.set_disabled(False)
                b.set_indicator(False if self.cur_mode == "cata" else None)
                b.set_text(f"{i + 1}    {self.choices[i]}")
                b.grid()
            else:
                b.grid_remove()

    def _load_card(self, slide):
        item = self.deck[self.idx]
        self.answered = self.hinted = self.last_wrong = False
        self.pending, self.picked = None, None
        self.removed, self.selected = set(), set()
        self.cur_mode = self._prepare(item, self.card_modes[self.idx])
        self._show_mode_ui()
        self.lbl_progress.config(
            text=f"Card {self.idx + 1} of {len(self.deck)}   \u00b7   "
                 f"{MODE_NAMES[self.cur_mode]}"
                 + ("   \u00b7   Review" if self.review else ""))
        self.progress.set(self.idx / len(self.deck))
        self._feedback("", "muted")
        self.answer_var.set("")
        self.entry.enable(True)
        self._update_score()
        self._set_buttons("ask")
        face = dict(tag=MODE_TAG[self.cur_mode],
                    category=item["category"].upper(), text=item["question"],
                    sub="", hint="", question="",
                    extra=(f"Answer: {self.tf_proposed}"
                           if self.cur_mode == "tf" else ""),
                    fill="card", outline="card_border")
        self.card.sx = self.card.pop = 1.0
        if slide:
            self.busy = True
            w = max(self.card.winfo_width(), 400)
            self.card.offset = w
            self.card.set_face(**face)
            SFX.play("swish")
            self.card.slide(w, 0, 360, done=self._ready)
        else:
            self.card.offset = 0
            self.card.set_face(**face)
            self._ready()

    def _ready(self):
        self.busy = False
        if self.round_start is None:                # the clock starts here
            now = time.perf_counter()
            self.round_start, self.paused_total = now, 0.0
            if self.paused:
                self.pause_started = now
        if not self.paused:
            if self.cur_mode == "id":
                self.entry.focus_field()
            else:
                self.focus_set()
        self._tick()

    # ---- player actions -------------------------------------------------------
    def hint(self):
        if (self.busy or self.paused or self.answered or self.hinted
                or self.cur_mode == "tf"):
            return
        self.hinted = True
        SFX.play("hint")
        if self.cur_mode == "mc":                 # cross out one wrong choice
            wrong = [i for i in range(len(self.choices))
                     if i != self.correct_idx and i not in self.removed]
            if len(wrong) >= 2:
                i = random.choice(wrong)
                self.removed.add(i)
                self.mc_btns[i].set_text("")
                self.mc_btns[i].set_disabled(True)
        elif self.cur_mode == "cata":             # say how many to tick
            n = len(self.right_idx)
            self.card.set_face(
                hint=f"Select {n} answer" + ("" if n == 1 else "s"))
        else:
            self.card.set_face(hint=make_hint(self.deck[self.idx]["answer"]))
            self.entry.focus_field()
        self.btn_hint.set_disabled(True)

    def choose(self, i):
        """Click / key on one of the option buttons."""
        if (self.busy or self.paused or self.answered
                or self.cur_mode == "id" or i >= len(self.choices)
                or i in self.removed):
            return
        if self.cur_mode == "cata":               # tick / untick
            self.selected ^= {i}
            on = i in self.selected
            self.mc_btns[i].set_kind("primary" if on else "secondary")
            self.mc_btns[i].set_indicator(on)
            return
        self.picked = i
        self._grade(self.choices[i], skipped=False)

    def submit(self):
        if (self.busy or self.paused or self.answered
                or self.cur_mode not in ("id", "cata")):
            return
        if self.cur_mode == "cata":
            if self.selected:
                self._grade("", skipped=False)
            return
        ans = self.entry.get().strip()
        if ans:
            self._grade(ans, skipped=False)

    def skip(self):
        if not (self.busy or self.paused or self.answered):
            self._grade("", skipped=True)

    def _color_choices(self):
        for i, b in enumerate(self.mc_btns[:len(self.choices)]):
            if self.cur_mode == "cata":
                if i in self.right_idx:
                    b.set_kind("good")
                elif i in self.selected:
                    b.set_kind("bad")
                b.set_indicator(i in self.selected)
            else:
                if i == self.correct_idx:
                    b.set_kind("good")
                elif i == self.picked:
                    b.set_kind("bad")
            b.set_disabled(True)

    def _grade(self, ans, skipped):
        item, mode = self.deck[self.idx], self.cur_mode
        self.answered, self.busy = True, True
        self.played += 1
        if mode in ("mc", "tf"):
            ok = (not skipped) and self.picked == self.correct_idx
            self._color_choices()
        elif mode == "cata":
            ok = (not skipped) and self.selected == self.right_idx
            self._color_choices()
        else:
            self.entry.enable(False)
            ok = (not skipped) and is_correct(ans, item["answer"])
        self.prev_streak = self.streak
        if ok:
            self.streak += 1
            self.best = max(self.best, self.streak)
            self.correct += 1
        else:
            self.streak = 0
            self.missed.append((item, mode))
        self.last_wrong = not ok
        self.pending = (item["row"], ok, item["type"])
        self._set_buttons("none")

        # the back of the card: the question again, then the answer
        if mode == "cata":
            main = "\n".join("\u2022  " + self.choices[i]
                             for i in sorted(self.right_idx))
            said = ", ".join(self.choices[i]
                             for i in sorted(self.selected)) or "nothing"
            sub = f"You selected: {said}"
        elif mode == "tf":
            main = item["answer"]
            sub = ("The statement was "
                   + ("TRUE" if self.correct_idx == 0 else "FALSE"))
        elif mode == "mc":
            main, sub = item["answer"], ""
        else:
            main, sub = item["answer"], f"You answered: {ans}"
        if ok:
            fill, outline = "good_bg", "good"
        elif skipped:
            fill, outline, sub = "neutral_bg", "accent", "You skipped this one"
        else:
            fill, outline = "bad_bg", "bad"
        back = dict(tag="ANSWER", question=item["question"], text=main,
                    sub=sub, extra="", hint="", fill=fill, outline=outline)
        SFX.play("flip")
        self.card.flip(back, done=lambda: self._after_flip(ok, skipped))

    def _after_flip(self, ok, skipped):
        self.busy = False
        if ok:
            SFX.play(f"correct{min(self.streak - 1, 5)}")
            self._feedback(self._msg("correct"), "good")
            self.card.pulse()
        else:
            if not skipped:
                SFX.play("wrong")
            self._feedback(self._msg("skip" if skipped else "wrong"), "bad")
            self.card.shake()
        self._update_score()
        self._set_buttons("review_wrong" if (not ok and not skipped
                                             and self.cur_mode == "id")
                          else "review_ok")

    def override(self):
        """Player says a typed answer should have counted."""
        if self.busy or self.paused or not (self.answered and self.last_wrong
                                            and self.pending):
            return
        self.last_wrong = False
        cur = self.deck[self.idx]
        self.missed = [m for m in self.missed if m[0] is not cur]
        self.streak = self.prev_streak + 1
        self.best = max(self.best, self.streak)
        self.correct += 1
        self.pending = (self.pending[0], True, self.pending[2])
        SFX.play(f"correct{min(self.streak - 1, 5)}")
        self.card.set_face(fill="good_bg", outline="good")
        self.card.pulse()
        self._feedback(self._msg("correct"), "good")
        self._update_score()
        self.btn_override.grid_remove()

    def _msg(self, key, pct=None):
        """A random message from the Messages sheet ({name}, {percent} work)."""
        if pct is None:
            pct = round(100 * self.correct / max(1, self.played))
        return (random.choice(self.msgs[key]).replace("{name}", self.player)
                .replace("{percent}", str(pct)))

    def _commit_pending(self):
        if self.pending:
            self.store.record_result(*self.pending)
            self.pending = None

    def next_card(self):
        if self.busy or self.paused or not self.answered:
            return
        self._commit_pending()
        self.idx += 1
        if self.idx >= len(self.deck):
            self.progress.set(1.0)
            self.finish()
            return
        self.busy = True
        self._set_buttons("none")
        w = max(self.card.winfo_width(), 400)
        self.card.slide(0, -w, 220, done=lambda: self._load_card(slide=True),
                        ease=ease_in)

    def end_round(self):
        if self.busy or self.paused or self.mode != "game":
            return
        self._pause(cover=False)                # stop the clock for the dialog
        if self.played:
            msg = (f"You've answered {self.played} of {len(self.deck)} "
                   "questions.\nEnd the round now and see your results?")
        else:
            msg = ("You haven't answered anything yet, so nothing will be "
                   "saved.\nEnd the round?")
        yes = self.confirm("End round?", msg, "End round", "Keep playing")
        self._resume(sound=False)
        if yes:
            if self.answered:
                self._commit_pending()
            self.finish()

    def finish(self):
        if self._timer_job is not None:
            self.after_cancel(self._timer_job)
            self._timer_job = None
        self.busy = False
        self.total_time = self._elapsed()
        if self.played == 0:
            self.show("menu")
            return
        pct = round(100 * self.correct / self.played)
        passing = self.msgs["passing"]
        passed = pct >= passing
        if not self.review:                     # review rounds stay off the board
            self.store.add_score(
                self.player, self.cat_label, self.played, self.correct,
                self.total_time, self.best, mode=self.mode_label,
                result="Passed" if passed else "Failed")
        saved = self.save()
        title = self._msg("pass" if passed else "fail", pct)
        self.res_title.config(text=title, font=(
            FONT, 28 if len(title) <= 40 else 22 if len(title) <= 80 else 17,
            "bold"))
        self.res_badge_key = "good" if passed else "warn"
        self.res_badge.config(
            text=("PASSED" if passed else "KEEP PRACTISING")
                 + f"   \u00b7   passing score {passing:g}%",
            fg=self.theme[self.res_badge_key])
        extra = ("Review round - not added to the leaderboard"
                 if self.review else self.mode_label)
        self.res_detail.config(
            text=f"{pct}% correct   \u00b7   Time {self._fmt_time(self.total_time)}\n"
                 f"Best streak: {self.best}   \u00b7   {extra}")
        self.res_saved.config(
            text=("Card stats saved to Excel." if self.review
                  else "Results saved to Excel.") if saved
            else "Results were NOT saved (file locked).")
        if self.missed:
            self.btn_review.set_text(f"Review mistakes ({len(self.missed)})")
            self.btn_review.grid()
        else:
            self.btn_review.grid_remove()
        self.show("results")
        SFX.play("finish" if passed else "done")
        right, asked = self.correct, self.played
        animate(self, 900,
                lambda p: self.res_score.config(
                    text=f"{round(right * p)} / {asked}"),
                ease=ease_out)

    # ------------------------------------------------------------------
    # results screen
    # ------------------------------------------------------------------
    def _build_results(self, f):
        f.columnconfigure(0, weight=1)
        self.res_title = self.lbl(f, "", 28, True, fg="accent")
        self.res_title.configure(wraplength=720, justify="center")
        self.res_title.grid(row=0, pady=(30, 4), padx=40)
        self.res_badge = tk.Label(f, bd=0, font=(FONT, 12, "bold"))
        self.reg(lambda t: self.res_badge.configure(
            bg=t["bg"], fg=t[self.res_badge_key]))
        self.res_badge.grid(row=1, pady=(0, 2))
        self.res_score = self.lbl(f, "", 60, True, fg="accent")
        self.res_score.grid(row=2, pady=4)
        self.lbl(f, "CORRECT", 10, True, fg="muted").grid(row=3)
        self.res_detail = self.lbl(f, "", 13)
        self.res_detail.grid(row=4, pady=(8, 4))
        self.res_saved = self.lbl(f, "", 10, fg="muted")
        self.res_saved.grid(row=5, pady=(2, 18))
        row = self.mkframe(f)
        row.grid(row=6)
        mk = lambda text, cmd, kind, w: self.reg_w(RoundedButton(
            row, text, cmd, self.theme, kind=kind, width=w, height=46,
            size=12))
        mk("Play again", self.start_game, "primary", 150).grid(
            row=0, column=0, padx=5)
        self.btn_review = mk("Review mistakes", self.start_review,
                             "secondary", 210)
        self.btn_review.grid(row=0, column=1, padx=5)
        mk("Leaderboard", self.show_leaderboard, "secondary", 140).grid(
            row=0, column=2, padx=5)
        mk("Menu", lambda: self.show("menu"), "secondary", 100).grid(
            row=0, column=3, padx=5)

    # ------------------------------------------------------------------
    # pop-up windows
    # ------------------------------------------------------------------
    def _table(self, title, headers, widths, rows, rank_col=False, footer=()):
        t = self.theme
        total = sum(widths) + 28
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry(f"{total + 40}x480")
        win.configure(bg=t["bg"])
        win.transient(self)
        if footer:                              # bottom bar first, so it stays
            bar = tk.Frame(win, bg=t["bg"])
            bar.pack(side="bottom", fill="x", padx=20, pady=(0, 12))
            for text, cmd in footer:
                RoundedButton(bar, text, lambda c=cmd: c(win), t, width=180,
                              height=38, size=10).pack(side="right")
        cv = tk.Canvas(win, bg=t["bg"], bd=0, highlightthickness=0)
        cv.pack(fill="both", expand=True)

        def clip(text, w):
            n = max(3, int(w / 7.5))
            text = str(text)
            return text if len(text) <= n else text[:n - 1] + "\u2026"

        x0, y = 20, 18
        x = x0 + 16
        for h, w in zip(headers, widths):
            cv.create_text(x, y + 10, text=h.upper(), anchor="w",
                           fill=t["muted"], font=(FONT, 9, "bold"))
            x += w
        y += 32
        for i, r in enumerate(rows):
            cv.create_polygon(rr_points(x0, y, x0 + total, y + 40, 16),
                              smooth=True, fill=t["card"],
                              outline=t["card_border"])
            x = x0 + 16
            for j, (val, w) in enumerate(zip(r, widths)):
                top = rank_col and j in (0, 2) and i < 3
                cv.create_text(x, y + 20, text=clip(val, w), anchor="w",
                               fill=t["accent"] if top else t["text"],
                               font=(FONT, 11, "bold" if top else "normal"))
                x += w
            y += 48
        cv.configure(scrollregion=(0, 0, total + 40, y + 16))
        wheel = lambda e: cv.yview_scroll(int(-e.delta / 120), "units")
        cv.bind("<MouseWheel>", wheel)
        win.bind("<MouseWheel>", wheel)
        return win

    def show_leaderboard(self):
        rows = self.store.leaderboard()
        if not rows:
            messagebox.showinfo("Leaderboard", "No scores yet - play a game!")
            return
        data = []
        for i, (p, c, q, secs, mode, cat, d, _res) in enumerate(rows, start=1):
            data.append((i, p, f"{c} / {q}",
                         f"{round(100 * c / q)}%" if q else "-",
                         "-" if secs is None else self._fmt_time(secs),
                         mode, cat, d))
        self._table("Leaderboard  (latest 20 games)",
                    ("#", "Player", "Correct", "%", "Time", "Mode",
                     "Category", "Date"),
                    (36, 100, 74, 52, 60, 160, 100, 120), data, rank_col=True,
                    footer=[("Reset leaderboard", self._reset_leaderboard)])

    def _reset_leaderboard(self, win):
        if self.confirm("Reset leaderboard?",
                        "This permanently erases every saved score.\n"
                        "It can't be undone.", "Reset", "Cancel", parent=win):
            self.store.clear_scores()
            self.save()
            win.destroy()
            messagebox.showinfo("Leaderboard", "The leaderboard was reset.")

    def show_hardest(self):
        items = [c for c in self.store.get_cards() + self.store.get_cata()
                 if c["seen"]]
        if not items:
            messagebox.showinfo("Hardest cards", "No card stats yet.")
            return
        items.sort(key=lambda c: c["correct"] / c["seen"])
        data = [(c["question"],
                 c["answer"] if c["type"] == "card" else ", ".join(c["right"]),
                 f"{c['correct']}/{c['seen']}",
                 f"{round(100 * c['correct'] / c['seen'])}%")
                for c in items[:25]]
        self._table("Hardest cards",
                    ("Question", "Answer", "Right / seen", "Accuracy"),
                    (310, 150, 100, 80), data)

    def add_card_dialog(self):
        t = self.theme
        win = tk.Toplevel(self)
        win.title("Add card")
        win.configure(bg=t["bg"], padx=26, pady=20)
        win.transient(self)
        win.grab_set()
        cat_var, q_var, a_var, w_var = (tk.StringVar(value="General"),
                                        tk.StringVar(), tk.StringVar(),
                                        tk.StringVar())

        def label(text, row):
            tk.Label(win, text=text, bg=t["bg"], fg=t["muted"], bd=0,
                     font=(FONT, 11)).grid(row=row, column=0, sticky="e",
                                           padx=10, pady=8)
        label("Category", 0)
        RoundedEntry(win, t, cat_var, width=290, height=42, size=12).grid(
            row=0, column=1, sticky="w", pady=8)
        RoundedButton(win, "\u25be", lambda: post_menu(
            win, t, self._categories(), cat_var.set), t, width=44,
            height=42, size=12).grid(row=0, column=2, padx=(6, 0))
        label("Question", 1)
        q_entry = RoundedEntry(win, t, q_var, width=340, height=42, size=12)
        q_entry.grid(row=1, column=1, columnspan=2, sticky="w", pady=8)
        label("Answer", 2)
        RoundedEntry(win, t, a_var, width=340, height=42, size=12).grid(
            row=2, column=1, columnspan=2, sticky="w", pady=8)
        label("Wrong choices", 3)
        RoundedEntry(win, t, w_var, width=340, height=42, size=12).grid(
            row=3, column=1, columnspan=2, sticky="w", pady=(8, 2))
        tk.Label(win, text="Optional - separate with ;   (used in Multiple "
                           "choice and True or False)",
                 bg=t["bg"], fg=t["muted"], bd=0, font=(FONT, 9)).grid(
            row=4, column=1, columnspan=2, sticky="w", padx=4)
        q_entry.focus_field()

        def save():
            q, a = q_var.get().strip(), a_var.get().strip()
            if not q or not a:
                messagebox.showwarning("Missing info",
                                       "Question and answer are required.",
                                       parent=win)
                return
            self.store.add_card(cat_var.get().strip() or "General", q, a,
                                Store._split(w_var.get()))
            self.save()
            self.refresh_menu()
            SFX.play("pop")
            win.destroy()

        row = tk.Frame(win, bg=t["bg"])
        row.grid(row=5, column=0, columnspan=3, pady=(14, 0))
        RoundedButton(row, "Save card", save, t, kind="primary", width=150,
                      height=44, size=12).grid(row=0, column=0, padx=5)
        RoundedButton(row, "Cancel", win.destroy, t, width=110, height=44,
                      size=12).grid(row=0, column=1, padx=5)
        win.bind("<Return>", lambda e: save())

    # ------------------------------------------------------------------
    # confirmation dialog, keyboard and window events
    # ------------------------------------------------------------------
    def confirm(self, title, message, yes="Yes", no="Cancel", parent=None):
        """A themed yes/no dialog. Returns True if the player confirms."""
        t, parent = self.theme, (parent or self)
        win = tk.Toplevel(self)
        win.title(title)
        win.configure(bg=t["bg"], padx=34, pady=26)
        win.transient(parent)
        win.resizable(False, False)
        result = {"yes": False}

        def done(v):
            result["yes"] = v
            win.destroy()
        tk.Label(win, text=title, bg=t["bg"], fg=t["accent"], bd=0,
                 font=(FONT, 17, "bold")).pack()
        tk.Label(win, text=message, bg=t["bg"], fg=t["text"], bd=0,
                 font=(FONT, 11), wraplength=380, justify="center").pack(
            pady=(10, 20))
        row = tk.Frame(win, bg=t["bg"])
        row.pack()
        RoundedButton(row, yes, lambda: done(True), t, kind="primary",
                      width=140, height=44, size=12).pack(side="left", padx=6)
        RoundedButton(row, no, lambda: done(False), t, width=140, height=44,
                      size=12).pack(side="left", padx=6)
        win.protocol("WM_DELETE_WINDOW", lambda: done(False))
        win.bind("<Return>", lambda e: done(True))
        win.bind("<Escape>", lambda e: done(False))
        win.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_reqwidth()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_reqheight()) // 3
        win.geometry(f"+{max(0, x)}+{max(0, y)}")
        win.grab_set()
        win.focus_force()
        self.wait_window(win)
        return result["yes"]

    def on_key(self, event):
        """1-6 pick / tick an option; T and F answer True or False."""
        if (self.mode != "game" or self.paused or self.cur_mode == "id"
                or not event.char):
            return
        ch = event.char.lower()
        if self.cur_mode == "tf" and ch in ("t", "f"):
            self.choose(0 if ch == "t" else 1)
        elif ch.isdigit() and ch != "0":
            self.choose(int(ch) - 1)

    def on_enter(self, _event=None):
        if self.mode == "game" and self.paused:
            self._resume()
            return
        if self.busy:
            return
        if self.mode == "menu":
            self.start_game()
        elif self.mode == "game":
            if self.answered:
                self.next_card()
            else:
                self.submit()
        elif self.mode == "results":
            self.start_game()

    def quit_app(self):
        """Quit button and the window's X: always ask first."""
        hold = self.mode == "game" and not self.paused
        if hold:
            self._pause(cover=False)
        msg = "Are you sure you want to quit?"
        if self.mode == "game":
            msg += "\nThe current round won't be added to the leaderboard."
        if self.confirm("Quit Flash Cards?", msg, "Quit", "Stay"):
            if self.mode == "game":
                self._commit_pending()
            self._save_prefs(flush=(self.mode != "game"))
            if self.mode == "game":
                self.save()
            self.destroy()
        elif hold:
            self._resume(sound=False)


def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else app_dir() / DEFAULT_FILE
    try:
        store = Store(path)
    except (SystemExit, Exception) as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Can't open deck", str(exc) or repr(exc))
        return
    App(store).mainloop()


if __name__ == "__main__":
    main()
