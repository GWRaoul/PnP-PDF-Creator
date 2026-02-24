# -*- coding: utf-8 -*-
"""
PnP PDF Creator
macOS writable paths patch (Documents-based)
Copyright (c) 2026 Raoul Schaupp

This software is provided free of charge.

Permission is granted to use, copy, and redistribute this software,
including the compiled executable, for private and non-commercial purposes,
provided that this notice is included unchanged.

Redistribution of this executable in unmodified form is explicitly permitted,
free of charge, including via third-party platforms such as itch.io.

------------------------------------------------------------
Third-Party Software and Licenses
------------------------------------------------------------

This product includes third-party software components:

- Python (Python Software Foundation License, PSF)
- PyInstaller (GPLv2 with bootloader exception)
- ReportLab (BSD-style license)
- Pillow (HPND license)
- Questionary (MIT License)
- Rich (MIT License)

All third-party components are used in accordance with their respective licenses.
Their original copyright notices remain with their respective authors.

------------------------------------------------------------
Disclaimer
------------------------------------------------------------

This software is provided "as is", without warranty of any kind,
express or implied, including but not limited to the warranties of
merchantability, fitness for a particular purpose, and non-infringement.
In no event shall the author be liable for any claim, damages, or other liability.
"""

import re
import os
import codecs
import textwrap
import tempfile
import hashlib
import sys
import configparser
import platform
import argparse
import io
from os.path import expanduser
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter, landscape
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import Color, HexColor, black
from reportlab.lib import colors

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    import questionary  # optionale Komfort-Listenprompts
except Exception:
    questionary = None
try:
    # Rich erzwingen (auch in PyInstaller-EXE ohne "volles" Terminal)
    from rich.console import Console
    from rich.panel import Panel
    from rich.table import Table
    from rich.progress import Progress, BarColumn, TextColumn, TimeRemainingColumn
    _FORCE_RICH = True
except Exception:
    Console = None  # type: ignore
    Panel = None    # type: ignore
    Table = None    # type: ignore
    Progress = None # type: ignore
    BarColumn = TextColumn = TimeRemainingColumn = None  # type: ignore
    _FORCE_RICH = False

# ---------------------------------------------------------------------------
# EXE/Terminal Kompatibilität erzwingen (Farben, Cursor, UTF-8)
# ---------------------------------------------------------------------------
# prompt_toolkit/questionary: volle Terminal-Features erzwingen
os.environ.setdefault("PROMPT_TOOLKIT_COLOR_DEPTH", "DEPTH_24_BIT")
os.environ.setdefault("PROMPT_TOOLKIT_FORCE_TERMINAL", "1")

try:
    if hasattr(sys.stdout, "buffer"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "buffer"):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
except Exception:
    pass

# Rich-Konsole initialisieren oder auf None setzen
console = Console(force_terminal=True, color_system="auto") if _FORCE_RICH and Console else None
rprint = (console.print if console else None)

# =========================================================
# Global "already paused" flag (prevents double Enter)
# =========================================================
_PAUSE_ALREADY_SHOWN = False

# =========================================================
# Script version / debug
# =========================================================
SCRIPT_VERSION = 'V1.3-2026-02-24'
DEBUG_PREPROCESS = False  # set True to print per-image crop/resize diagnostics

# =========================================================
# Quality presets (Cards only)
# =========================================================
DEFAULT_QUALITY = "high"  # default on Enter

QUALITY_PRESETS = {
    "high":   {"dpi": 300, "jpeg_quality": 90},
    "medium": {"dpi": 200, "jpeg_quality": 80},
    "low":    {"dpi": 120, "jpeg_quality": 65},
}

# =========================================================
# Global config
# =========================================================
SUPPORTED_EXT = {".png", ".jpg", ".jpeg"}
# Default basename for a shared card back image (configurable via INI [assets]/cardback_name)
DEFAULT_CARDBACK_BASENAME = 'cardback'
CARDBACK_BASENAME = DEFAULT_CARDBACK_BASENAME
DEFAULT_LOGO_BASENAME = 'logo'
LOGO_BASENAME = DEFAULT_LOGO_BASENAME
DEFAULT_RULEBOOK_BASENAME = 'rulebook'
RULEBOOK_BASENAME = DEFAULT_RULEBOOK_BASENAME
DEFAULT_RULEBOOK_ROTATE_MODE = 'auto'  # auto|off|force_landscape|force_portrait
RULEBOOK_ROTATE_MODE = DEFAULT_RULEBOOK_ROTATE_MODE

# Card format templates (fixed template DPI; bleed is 1/8" per side)
TEMPLATE_DPI = 300
BLEED_IN_PER_SIDE = 0.125  # 1/8"
MM_PER_INCH = 25.4

# Zentrale Laufzeit-STATE (vermeidet global/Annotation-Konflikte)
STATE = {
    "current_format": None  # wird in apply_card_format(fmt) gesetzt
}

# Predefined card formats (UI order is frozen).
# UI line format: <Name> (<w> x <h> mm) [n]
CARD_FORMATS = [
    {'id': 1, 'name': 'Poker', 'w_mm': 63.5, 'h_mm': 88.9},
    {'id': 2, 'name': 'Euro', 'w_mm': 59.0, 'h_mm': 92.0},
    {'id': 3, 'name': 'Mini Euro', 'w_mm': 44.0, 'h_mm': 68.0},
    {'id': 4, 'name': 'American', 'w_mm': 56.0, 'h_mm': 87.0},
    {'id': 5, 'name': 'Mini American', 'w_mm': 41.0, 'h_mm': 63.0},
]

def _fmt_to_inner_px(w_mm: float, h_mm: float) -> tuple[int, int]:
    # Inner/trim pixels at TEMPLATE_DPI
    w_in = w_mm / MM_PER_INCH
    h_in = h_mm / MM_PER_INCH
    return int(round(w_in * TEMPLATE_DPI)), int(round(h_in * TEMPLATE_DPI))

def _px_to_mm(px: float) -> float:
    # TEMPLATE_DPI = 300, MM_PER_INCH = 25.4
    return (px / TEMPLATE_DPI) * MM_PER_INCH

def apply_card_format(fmt: dict) -> None:
    """Apply selected card format to global geometry variables (format-rein).

    We keep the existing pipeline by updating the global INNER/BLEED/PT constants.
    Bleed is always 1/8" per side at TEMPLATE_DPI => 37.5 px; split as 37/38 like the Poker template.
    """
    global POKER_W_PT, POKER_H_PT
    global BLEED_W_PX, BLEED_H_PX, INNER_W_PX, INNER_H_PX


    w_mm = float(fmt['w_mm'])
    h_mm = float(fmt['h_mm'])
    w_in = w_mm / MM_PER_INCH
    h_in = h_mm / MM_PER_INCH

    # Physical size in PDF points (trim size)
    POKER_W_PT = w_in * 72.0
    POKER_H_PT = h_in * 72.0

    # Template pixel sizes
    iw, ih = _fmt_to_inner_px(w_mm, h_mm)
    INNER_W_PX = iw
    INNER_H_PX = ih
    BLEED_W_PX = iw + BLEED_LEFT_TOP_PX + BLEED_RIGHT_BOTTOM_PX
    BLEED_H_PX = ih + BLEED_LEFT_TOP_PX + BLEED_RIGHT_BOTTOM_PX
    STATE["current_format"] = fmt

def is_mini_format(fmt: dict) -> bool:
    """True für Mini Euro / Mini American (Name enthält 'Mini')."""
    return 'mini' in str(fmt.get('name', '')).lower()

def prompt_card_format() -> dict:
    """Prompt user for card format. Enter selects Poker (id=1)."""
    print(t('choose_card_format'))
    # Einmal über alle Formate iterieren und ausgeben
    for f in CARD_FORMATS:
        w = float(f['w_mm'])
        h = float(f['h_mm'])
        # Custom (INI) -> 1 Nachkommastelle, sonst kompakt
        if f.get('src') == 'ini':
            w_str = _mm_str_custom(w)
            h_str = _mm_str_custom(h)
        else:
            w_str = _mm_str(w)
            h_str = _mm_str(h)
        print(f"{f['name']} ({_mm_str(w)} x {_mm_str(h)} mm) [{f['id']}]")

    # Auswahl vom Nutzer einlesen (unverändert)
    while True:
        raw = input(t('choose_card_format_prompt')).strip()
        if raw == '':
            choice = 1
        else:
            try:
                choice = int(raw)
            except Exception:
                print(t('invalid_card_format'))
                continue
        fmt = next((f for f in CARD_FORMATS if f['id'] == choice), None)
        if fmt is None:
            print(t('invalid_card_format'))
            continue
        return fmt

def _mm_str(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else str(v).rstrip('0').rstrip('.')

def _mm_str_custom(v: float) -> str:
    # Immer 1 Nachkommastelle
    return f"{v:.1f}"

def print_selected_format_info(fmt: dict) -> None:
    """Always print selected format and expected image sizes (localized)."""
    if fmt.get('src') == 'ini':
        w = _mm_str_custom(float(fmt['w_mm']))
        h = _mm_str_custom(float(fmt['h_mm']))
    else:
        w = _mm_str(float(fmt['w_mm']))
        h = _mm_str(float(fmt['h_mm']))
    # Inner/bleed pixels are based on current globals (already applied)
    iw, ih = INNER_W_PX, INNER_H_PX
    bw, bh = BLEED_W_PX, BLEED_H_PX
    print('')
    print(t('format_info_header', name=fmt['name'], w=w, h=h))
    print(t('format_info_sizes', iw=iw, ih=ih, bw=bw, bh=bh))
    print(t('format_info_note'))
    print('')


# Poker (inner) size in PDF points (2.5" x 3.5")
POKER_W_PT = 2.5 * 72
POKER_H_PT = 3.5 * 72

# Logo constraints (placement scaling only, NO compression)
LOGO_MAX_W = 200.0
LOGO_MAX_H = 15.0
LOGO_GAP_TO_GRID = 2.0

# Bottom line
COPY_MAX_CHARS = 150

# 3x3 + Gutterfold cut marks (standard)
# These are overridable via INI (section [cutmarks]).
CUTMARK_LEN_PT_STD = 5.0
CUTMARK_LINE_PT_STD = 1.0

# 2x3 marks (outer only, cut to poker area inside bleed image)
CUTMARK_LEN_PT_BLEED = 20.0
CUTMARK_LINE_PT_BLEED = 1.0
# 2x3 card image geometry (pixels of the source image)
BLEED_W_PX = 825
BLEED_H_PX = 1125
BLEED_LEFT_TOP_PX = 37
BLEED_RIGHT_BOTTOM_PX = 38
INNER_W_PX = BLEED_W_PX - BLEED_LEFT_TOP_PX - BLEED_RIGHT_BOTTOM_PX  # 750
INNER_H_PX = BLEED_H_PX - BLEED_LEFT_TOP_PX - BLEED_RIGHT_BOTTOM_PX  # 1050

# =========================================================
# Dünnen Außen-Bleed nur für Standard & Gutterfold
# =========================================================
# Pixelangabe bezieht sich auf die 300-dpi-Bleed-Canvas. Wird beim Zeichnen
# sauber auf Punkte skaliert, sodass die Innenfläche exakt 2.5"x3.5" bleibt.
OUTER_BLEED_KEEP_PX = 15

# =========================================================
# Gutterfold layout config (NEW: 2 rows x 4 cols, horizontal fold)
# =========================================================

GF_FOLD_GUTTER_PT = 12.0   # Abstand zwischen oberer und unterer Reihe (Falzbereich)
GF_COL_GAP_PT = 0.0        # <-- BÜNDIG (keine Lücke zwischen Spalten)

GF_DRAW_FOLD_LINE = True
GF_FOLD_LINE_WIDTH = 0.8
GF_FOLD_LINE_DASH = (3, 3)

# Placement of page number and version number
LEFT_MARGIN = 20.0 # Version number
RIGHT_MARGIN = 20.0 # Page number
BOTTOM_Y = 10.0 # Placement from bottom page
BOTTOM_Y_LETTER_3X3 = 6.0  # nur 3x3 im Letter-Format: tiefer setzen

# =========================================================
# Druckfreier Rand + Reserven (NEU)
# =========================================================
def cm_to_pt(cm: float) -> float:
    return cm * 72.0 / 2.54

# 1 cm druckfreier Rand rundum
PRINT_SAFE_MARGIN_CM = 0.1
MARGINS_PT = {
    "left":   cm_to_pt(PRINT_SAFE_MARGIN_CM),
    "right":  cm_to_pt(PRINT_SAFE_MARGIN_CM),
    "top":    cm_to_pt(PRINT_SAFE_MARGIN_CM),
    "bottom": cm_to_pt(PRINT_SAFE_MARGIN_CM),
}

# Untere Reserve: verhindert, dass Karten die Fußzeile überdecken
# (mind. etwas über BOTTOM_Y sowie lang genug für Außenmarken)
BOTTOM_RESERVED_PT = max(BOTTOM_Y + 5.0, CUTMARK_LEN_PT_STD)

# ---------------------------------------------------------
# Reservierter Platz für Kopf/Fuß und Seiten (konfigurierbar)
# Symmetrisch als Default, damit die optische Zentrierung gewahrt bleibt.
# ---------------------------------------------------------
RESERVE_BOTTOM_PT = BOTTOM_RESERVED_PT
RESERVE_TOP_PT    = RESERVE_BOTTOM_PT
RESERVE_LEFT_PT   = 0.0
RESERVE_RIGHT_PT  = 0.0

# --- Backside offset (read from INI) ---
# Werte werden in mm in der INI gepflegt und hier als PDF-Punkte geführt.
# Positive X -> nach rechts; positive Y -> nach oben.
BACK_X_OFFSET_PT = 0.0
BACK_Y_OFFSET_PT = 0.0

def _mm_to_pt(mm: float) -> float:
    return (mm / 25.4) * 72.0

def draw_logo_in_header_band(c, logo_path, page_w, page_h, margins, header_h):
    lw, lh = fit_logo_with_constraints(logo_path, LOGO_MAX_W, LOGO_MAX_H)
    x = margins["left"] + RESERVE_LEFT_PT + (page_w - margins["left"] - margins["right"]
                                             - RESERVE_LEFT_PT - RESERVE_RIGHT_PT - lw)/2.0
    y = page_h - margins["top"] - header_h + (header_h - lh)/2.0  # mittig im Kopfband
    c.drawImage(ImageReader(str(logo_path)), x, y, width=lw, height=lh,
                preserveAspectRatio=True, mask="auto")


# =========================================================
# Language (UI) - EXE-safe persistence via INI next to EXE
# =========================================================

LANG = ""  # runtime value

# IMPORTANT: The *real* app name (Documents/<APP_NAME>/...)
APP_NAME = "PnP PDF Creator"

def get_app_dir() -> Path:
    """
    Directory where the EXE resides (PyInstaller) or script directory (normal python).
    For PyInstaller --onefile, sys.executable points to the actual .exe path.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def is_macos_app() -> bool:
    """
    True if running as a frozen PyInstaller app on macOS.
    This setup is subject to App Translocation read-only bundle paths.
    """
    return bool(getattr(sys, "frozen", False)) and sys.platform == "darwin"

def get_macos_documents_base_dir() -> Path:
    """
    Variant B (user-friendly):
    Use ~/Documents/PnP PDF Creator/ as writable base directory.
    """
    base = Path.home() / "Documents" / APP_NAME
    base.mkdir(parents=True, exist_ok=True)
    return base

def get_writable_base_dir() -> Path:
   """
    Windows/Linux: legacy (beside EXE/script).
    macOS App: write to Documents/<APP_NAME>/ to avoid App Translocation paths.
    """
    if is_macos_app():
        return get_macos_documents_base_dir()
    return get_app_dir()

def make_safe_name(name: str) -> str:
    """
    Sanitizes a UI-provided base name for use in a folder:
    - trims
    - replaces any non [A-Za-z0-9._-] with '_'
    - strips leading/trailing underscores
    """
    import re as _re
    if not name:
        return "output"
    s = _re.sub(r"[^A-Za-z0-9._-]+", "_", name.strip())
    return s.strip("_") or "output"

def build_generation_dir(out_base: str) -> Path:
    """
    Creates a per-run subfolder below the EXE/script directory with the pattern:
    <YYYYMMDD>_<HHMMSS>_<out_base-sanitized>
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    folder_name = f"{ts}_{make_safe_name(out_base)}"
    gen_dir = get_writable_base_dir() / folder_name
    gen_dir.mkdir(parents=True, exist_ok=True)
    return gen_dir

def get_ini_path() -> Path:
    return get_writable_base_dir() / "PnP_PDF_Creator.ini"

# =========================================================
# INI handling (UI language + cutmark settings)
# =========================================================

def load_config() -> configparser.ConfigParser:
    cp = configparser.ConfigParser()
    ini_path = get_ini_path()
    if ini_path.exists():
        try:
            cp.read(ini_path, encoding='utf-8')
        except Exception:
            # broken INI -> ignore
            pass
    return cp


def write_config(cp: configparser.ConfigParser) -> None:
    ini_path = get_ini_path()
    try:
        with ini_path.open('w', encoding='utf-8') as f:
            cp.write(f)
    except Exception as e:
        print(f"[WARN] Could not write INI next to EXE: {ini_path} ({e})")
        print("[WARN] If the EXE is in a protected folder (e.g. Program Files), move it to a writable folder.")


def _get_positive_float(cp: configparser.ConfigParser, section: str, option: str, fallback: float) -> float:
    try:
        v = cp.getfloat(section, option, fallback=fallback)
        if v <= 0:
            return fallback
        return float(v)
    except Exception:
        return fallback
        
def _get_outer_bleed_keep_px(cp: configparser.ConfigParser,
                             section: str,
                             option: str,
                             fallback: int) -> int:
    try:
        v = cp.getint(section, option, fallback=fallback)
    except Exception:
        return fallback

    if v < 0:
        return 0
    if v > 20:
        return 20
    return v       

def ensure_cutmark_defaults(cp: configparser.ConfigParser) -> bool:
    # Ensure [cutmarks] section exists with defaults. Returns True if cp was modified.
    changed = False
    if not cp.has_section('cutmarks'):
        cp.add_section('cutmarks')
        changed = True
    defaults = {
        'length_pt_standard': str(CUTMARK_LEN_PT_STD),
        'width_pt_standard': str(CUTMARK_LINE_PT_STD),
        'length_pt_bleed': str(CUTMARK_LEN_PT_BLEED),
        'width_pt_bleed': str(CUTMARK_LINE_PT_BLEED),
    }
    for k, v in defaults.items():
        if not cp.has_option('cutmarks', k):
            cp.set('cutmarks', k, v)
            changed = True
    return changed

def ensure_standard_and_gutterfold_defaults(cp: configparser.ConfigParser) -> bool:
    """
    Stellt sicher, dass die Sektion [standard_and_gutterfold]
    existiert und den Parameter outer_bleed_keep_px enthält.
    """
    changed = False

    if not cp.has_section("standard_and_gutterfold"):
        cp.add_section("standard_and_gutterfold")
        changed = True

    if not cp.has_option("standard_and_gutterfold", "outer_bleed_keep_px"):
        cp.set("standard_and_gutterfold", "outer_bleed_keep_px", str(OUTER_BLEED_KEEP_PX))
        changed = True

    return changed

def ensure_custom_format_defaults(cp: configparser.ConfigParser) -> bool:
    """
    Stellt [custom_format] mit sinnvollen Defaults bereit.
    Rückgabe: True, wenn cp geändert wurde.
    """
    changed = False
    if not cp.has_section('custom_format'):
        cp.add_section('custom_format')
        changed = True
    if not cp.has_option('custom_format', 'name'):
        cp.set('custom_format', 'name', 'Custom (INI)')
        changed = True
    # Default = Poker-Trim 750x1050 px, damit sofort nutzbar
    if not cp.has_option('custom_format', 'inner_w_px'):
        cp.set('custom_format', 'inner_w_px', '750')
        changed = True
    if not cp.has_option('custom_format', 'inner_h_px'):
        cp.set('custom_format', 'inner_h_px', '1050')
        changed = True
    return changed

def ensure_backside_offset_defaults(cp: configparser.ConfigParser) -> bool:
    """Sichert [backside_offset] + Defaults (x_offset/y_offset in mm)."""
    changed = False
    if not cp.has_section('backside_offset'):
        cp.add_section('backside_offset')
        changed = True
    if not cp.has_option('backside_offset', 'x_offset'):
        cp.set('backside_offset', 'x_offset', '0')
        changed = True
    if not cp.has_option('backside_offset', 'y_offset'):
        cp.set('backside_offset', 'y_offset', '0')
        changed = True
    return changed
 
def load_backside_offset_from_config(cp: configparser.ConfigParser) -> None:
    """Liest mm-Werte aus [backside_offset] und pflegt globale *_PT."""
    global BACK_X_OFFSET_PT, BACK_Y_OFFSET_PT
    try:
        x_mm = cp.getfloat('backside_offset', 'x_offset', fallback=0.0)
    except Exception:
        x_mm = 0.0
    try:
        y_mm = cp.getfloat('backside_offset', 'y_offset', fallback=0.0)
    except Exception:
        y_mm = 0.0
    BACK_X_OFFSET_PT = _mm_to_pt(float(x_mm))
    BACK_Y_OFFSET_PT = _mm_to_pt(float(y_mm))

def load_custom_format_from_config(cp: configparser.ConfigParser) -> Optional[dict]:
    """
    Liest [custom_format] und baut ein Format-Dict im Stil von CARD_FORMATS.
    Erwartet 'name', 'inner_w_px', 'inner_h_px' > 0.
    """
    try:
        name = cp.get('custom_format', 'name', fallback='').strip()
        w_px = cp.getint('custom_format', 'inner_w_px', fallback=0)
        h_px = cp.getint('custom_format', 'inner_h_px', fallback=0)
        if not name or w_px <= 0 or h_px <= 0:
            return None
        # mm aus px bei TEMPLATE_DPI
        w_mm = _px_to_mm(float(w_px))
        h_mm = _px_to_mm(float(h_px))
        # ID 6 reservieren
        return {'id': 6, 'name': name, 'w_mm': w_mm, 'h_mm': h_mm, 'src': 'ini'}
    except Exception:
        return None

def ensure_assets_defaults(cp: configparser.ConfigParser) -> bool:
    # Ensure [assets] section exists with defaults (e.g. shared cardback image name, logo name).
    changed = False
    if not cp.has_section('assets'):
        cp.add_section('assets')
        changed = True
    if not cp.has_option('assets', 'cardback_name'):
        cp.set('assets', 'cardback_name', DEFAULT_CARDBACK_BASENAME)
        changed = True
    if not cp.has_option('assets', 'logo_name'):
        cp.set('assets', 'logo_name', DEFAULT_LOGO_BASENAME)
        changed = True
    if not cp.has_option('assets', 'rulebook_name'):
        cp.set('assets', 'rulebook_name', DEFAULT_RULEBOOK_BASENAME)
        changed = True
    if not cp.has_option('assets', 'rulebook_rotate'):
        cp.set('assets', 'rulebook_rotate', DEFAULT_RULEBOOK_ROTATE_MODE)
        changed = True
    return changed

def load_assets_from_config(cp: configparser.ConfigParser) -> None:
    # Load asset settings from INI into global variables.
    global CARDBACK_BASENAME, LOGO_BASENAME, RULEBOOK_BASENAME, RULEBOOK_ROTATE_MODE
    # cardback
    name = cp.get('assets', 'cardback_name', fallback=DEFAULT_CARDBACK_BASENAME).strip()
    CARDBACK_BASENAME = name if name else DEFAULT_CARDBACK_BASENAME
    # logo
    logo_name = cp.get('assets', 'logo_name', fallback=DEFAULT_LOGO_BASENAME).strip()
    LOGO_BASENAME = logo_name if logo_name else DEFAULT_LOGO_BASENAME
    # rulebook
    rulebook_name = cp.get('assets', 'rulebook_name', fallback=DEFAULT_RULEBOOK_BASENAME).strip()
    RULEBOOK_BASENAME = rulebook_name if rulebook_name else DEFAULT_RULEBOOK_BASENAME
    # rulebook rotate mode
    rotate_mode = cp.get('assets', 'rulebook_rotate', fallback=DEFAULT_RULEBOOK_ROTATE_MODE).strip().lower()
    if rotate_mode not in ('auto', 'off', 'force_landscape', 'force_portrait'):
        rotate_mode = DEFAULT_RULEBOOK_ROTATE_MODE
    RULEBOOK_ROTATE_MODE = rotate_mode    

def load_cutmarks_from_config(cp: configparser.ConfigParser) -> None:
    # Load cutmark settings from INI into the global variables.
    global CUTMARK_LEN_PT_STD, CUTMARK_LINE_PT_STD, CUTMARK_LEN_PT_BLEED, CUTMARK_LINE_PT_BLEED, CUTMARK_COLOR, OUTER_BLEED_KEEP_PX
    CUTMARK_LEN_PT_STD = _get_positive_float(cp, 'cutmarks', 'length_pt_standard', CUTMARK_LEN_PT_STD)
    CUTMARK_LINE_PT_STD = _get_positive_float(cp, 'cutmarks', 'width_pt_standard', CUTMARK_LINE_PT_STD)
    CUTMARK_LEN_PT_BLEED = _get_positive_float(cp, 'cutmarks', 'length_pt_bleed', CUTMARK_LEN_PT_BLEED)
    CUTMARK_LINE_PT_BLEED = _get_positive_float(cp, 'cutmarks', 'width_pt_bleed', CUTMARK_LINE_PT_BLEED)
    CUTMARK_COLOR = cp.get('cutmarks', 'cutmark_color', fallback='#000000').strip()


def parse_color(value: str):
    """
    Akzeptiert Hex (#RRGGBB) oder benannte Farben ('red', 'blue', ...).
    Fällt bei Fehlern auf Schwarz zurück.
    """
    try:
        v = (value or "").strip()
        # lässt #RRGGBB / RRGGBB / benannte Farben zu
        col = colors.toColor(v)  # kann #hex, 'red', '0xRRGGBB', etc.
        return col
    except Exception:
        return black

def save_lang_to_ini(lang: str) -> None:
    cp = load_config()
    if not cp.has_section('ui'):
        cp.add_section('ui')
    cp.set('ui', 'lang', lang)
    # Ensure cutmark defaults exist so users can edit them
    ensure_cutmark_defaults(cp)
    # Ensure assets defaults exist so users can edit them
    ensure_assets_defaults(cp)    
    # --- SAFETY: Stelle sicher, dass cutmark_color wirklich gesetzt ist ---
    if not cp.has_option('cutmarks', 'cutmark_color'):
        cp.set('cutmarks', 'cutmark_color', '#000000')
    write_config(cp)

def prompt_language_if_needed():
    global LANG
    cp = load_config()
    changed = ensure_cutmark_defaults(cp)
    changed = ensure_assets_defaults(cp) or changed
    changed = ensure_custom_format_defaults(cp) or changed
    changed = ensure_standard_and_gutterfold_defaults(cp) or changed
    changed = ensure_backside_offset_defaults(cp) or changed
    
    # Optional: gleich laden & an CARD_FORMATS anhängen (am Ende der Liste)
    fmt6 = load_custom_format_from_config(cp)
    if fmt6:
        # Prüfen, ob ID 6 bereits existiert (Sicherheitsnetz)
        if not any(f.get('id') == 6 for f in CARD_FORMATS):
            CARD_FORMATS.append(fmt6)

    # 1) Try loading from INI first
    lang = cp.get('ui', 'lang', fallback='').strip().lower()
    if lang in ('de', 'en', 'fr', 'es', 'it'):
        LANG = lang
    else:
        # 2) Ask user once
        print('Please select language: de (Deutsch), en (English), fr (Français), es (Español), it (Italiano)')
        choice = input('Language [en]: ').strip().lower()
        if choice in ('de', 'en', 'fr', 'es', 'it'):
            LANG = choice
        else:
            LANG = 'en'
        if not cp.has_section('ui'):
            cp.add_section('ui')
        cp.set('ui', 'lang', LANG)
        changed = True
        print(f"Language saved to {get_ini_path()}: '{LANG}'")

    # Load cutmark settings into globals
    load_cutmarks_from_config(cp)
    # Load asset settings into globals
    load_assets_from_config(cp)
    # Standard & Gutterfold Bleed aus neuer Sektion laden
    global OUTER_BLEED_KEEP_PX
    OUTER_BLEED_KEEP_PX = _get_outer_bleed_keep_px(
        cp,
        "standard_and_gutterfold",
        "outer_bleed_keep_px",
        OUTER_BLEED_KEEP_PX
    )
    # Backside-Offset laden (mm -> pt)
    load_backside_offset_from_config(cp)
    # --- SAFETY: Stelle sicher, dass cutmark_color wirklich vorhanden ist ---
    if not cp.has_option('cutmarks', 'cutmark_color'):
        cp.set('cutmarks', 'cutmark_color', '#000000')
        changed = True
    # Persist INI if defaults were added or language was set
    if changed:
        write_config(cp)

I18N = {
    "de": {
        "choose_layout": "Layout wählen ({opts}) [All]: ",
        "startup_license": "Freie Software – siehe LICENSE.txt (und Lizenzhinweis im Header).",
        "header_welcome": "Willkommen",
        "no_cards_title": "Keine Karten gefunden",
        "invalid_layout": "Bitte eines der angebotenen Layouts eingeben.",
        "format_info_note": "'Standard' verwendet Innenbilder ohne Beschnitt. 'Bleed' benötigt Bilder mit Beschnitt. 'Gutterfold' erstellt ein Falzlayout mit vorderer/hinterer Seite.",
        "skip_2x5": "2x5 wird übersprungen: Kartenbilder haben keinen Bleed (mind. {minw}x{minh}) oder sind gemischt.",
        "format_info_header": "Ausgewähltes Kartenformat: {name} ({w} x {h} mm)",
        "format_info_sizes": "Erwartete Bildgrößen @300dpi: Innen {iw}x{ih} px, Bleed {bw}x{bh} px (Bleed = 1/8\" pro Seite)",
        "choose_card_format": "Kartenformat wählen (Zahl eingeben, Enter = Poker):",
        "choose_card_format_prompt": "Auswahl [1]: ",
        "invalid_card_format": "Ungültige Auswahl. Bitte eine Zahl aus der Liste eingeben.",
        "choose_format": "Papierformat waehlen (A4/Letter/Both) [Both]: ",
        "invalid_format": "Bitte 'A4', 'Letter' oder 'Both' eingeben.",
        "ask_folder": "Pfad zum Ordner mit PNG/JPG Kartenbildern: ",
        "invalid_folder": "Ungueltiger Pfad oder kein Ordner. Bitte erneut eingeben.",
        "invalid_folder_path": "Ungueltiger Pfad oder kein Ordner: {path}\nBitte erneut eingeben.",
        "ask_logo": "Optional: Pfad zur Logo-Datei (Enter = auto: logo.png/jpg im Kartenordner): ",
        "logo_invalid": "Logo-Pfad ungueltig oder Datei nicht gefunden. Es wird ohne Logo fortgefahren.",
        "ask_quality": "Qualitaet waehlen (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Ungueltige Qualitaet. Bitte Lossless, High, Medium oder Low eingeben.",
        "ask_copyright": "Freitext unten (optional, max. 150 Zeichen): ",
        "ask_version": "Versionsnummer eingeben (Enter = leer): ",
        "ask_out_base": "Ausgabedatei Basisname (ohne .pdf) [{default}]: ",
        "no_cards": (
            "Keine Karten gefunden im Ordner:\n{folder}\n\n"
            "Erwartet: Dateinamen enden auf 'a' oder 'b' (z.B. card01a.png / card01b.png) "
            "ODER enden auf '[face,<n>]' bzw. '[back,<n>]' (z.B. card01[face,001].png / card01[back,001].png)."
        ),
        "no_cards_examples": "Beispiele im Ordner: {files}{more}",
        "no_cards_no_images": "Keine Bilddateien (.png/.jpg/.jpeg) im Ordner gefunden.",
        "done": "Fertig! PDF erstellt: {path}",
        "skip_2x3": "2x3 wird übersprungen: Kartenbilder haben keinen Bleed (mind. {minw}x{minh}) oder sind gemischt.",
        "skip_gutterfold_no_backs": "Gutterfold wird übersprungen: Keine Rückseiten gefunden und keine Datei mit dem Namen '{name}' im Kartenordner.",
        "using_cardback": "Keine Rückseiten gefunden – verwende '{file}' als gemeinsame Kartenrückseite für alle Karten.",
        "skip_gutterfold_missing_backs": "Gutterfold wird übersprungen: Nicht für alle Vorderseiten wurde eine Rückseite gefunden. Fehlende Rückseiten für: {missing}",
        "skip_bleed_due_to_small": "Bleed wird nicht erzeugt: Mindestens eine Datei unterschreitet die Bleed-Mindestgröße {minw}x{minh} Pixel ({count} Datei(en)). Diese Dateien werden übersprungen: {files}",
        "warn_too_small_upscale": "Hinweis: Mindestens ein Kartenbild ist kleiner als {minw}x{minh} Pixel ({count} Datei(en)). Es wird auf diese Größe hochskaliert: {files}",
        "error_gutterfold_space": (
            "Zu wenig Platz für Gutterfold: verfügbar {avail:.1f} pt, "
            "benötigt {need:.1f} pt (Rand: {margin} cm; Kopf-Reserve: {top:.1f} pt; "
            "Fuß-Reserve: {bottom:.1f} pt)."   
        ),
        "rulebook_found": "Rulebook-Seiten gefunden: {files}",
        "rulebook_not_found": "Keine Rulebook-Seiten gefunden (Suche nach '{name}*').",
        "rulebook_will_prepend": "Gefundene Rulebook-Seiten werden vorne in das PDF eingefügt.",
        "logo_found": "Logo gefunden: {file}",
        "logo_not_found": "Kein Logo gefunden (gesucht nach '{name}'). Es wird ohne Logo fortgefahren.",
        "count_mismatch_warn": "Abweichende Anzahl bei '{base}': face={face} back={back} -> verwende face={use}",
        "using_pdfconfig": "Verwende pdfConfig.txt aus:\n{path}\nAlle UI-Eingaben werden übersprungen.",
        "config_title": "Konfiguration",
        "cfg_format": "Kartenformat: {id} ({name})",
        "cfg_layouts": "Layout(s): {layouts}",
        "cfg_paper": "Papierformat: {paper}",
        "cfg_quality": "Qualität: {quality}",
        "cfg_bottom_text": "Fußzeilentext: {text}",
        "cfg_version": "Version: {version}",
        "cfg_output_name": "Dateiname: {name}",
        "none": "— kein —",
        "exit_ok": "Fertig. Programmende.",
        "exit_err_header": "[FEHLER] Es ist ein unerwarteter Fehler aufgetreten:",
        "exit_press_enter": "Drücke Enter, um zu schließen…",
    },
    "en": {
        "choose_layout": "Choose layout ({opts}) [All]: ",
        "startup_license": "Free software – see LICENSE.txt (and header license notice).",
        "header_welcome": "Welcome",
        "no_cards_title": "No cards found",
        "invalid_layout": "Please enter one of the offered layouts.",
        "format_info_note": "'Standard' uses inner images (no bleed). 'Bleed' requires bleed images. 'Gutterfold' produces a fold layout with matching front/back alignment.",
        "skip_2x5": "Skipping 2x5: card images do not have bleed (min {minw}x{minh}) or are mixed.",
        "format_info_header": "Selected card format: {name} ({w} x {h} mm)",
        "format_info_sizes": "Expected image sizes @300dpi: inner {iw}x{ih} px, bleed {bw}x{bh} px (bleed = 1/8\" per side)",
        "choose_card_format": "Choose card format (enter number, Enter = Poker):",
        "choose_card_format_prompt": "Selection [1]: ",
        "invalid_card_format": "Invalid selection. Please enter a number from the list.",
        "choose_format": "Choose paper size (A4/Letter/Both) [Both]: ",
        "invalid_format": "Please enter 'A4', 'Letter' or 'Both'.",
        "ask_folder": "Path to folder with PNG/JPG card images: ",
        "invalid_folder": "Invalid path or not a folder. Please try again.",
        "invalid_folder_path": "Invalid path or not a folder: {path}\nPlease try again.",
        "ask_logo": "Optional: path to logo file (Enter = auto: logo.png/jpg in card folder): ",
        "logo_invalid": "Invalid logo path or file not found. Continuing without logo.",
        "ask_quality": "Choose quality (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Invalid quality. Please enter Lossless, High, Medium or Low.",
        "ask_copyright": "Bottom free text (optional, max 150 chars): ",
        "ask_version": "Enter version string (Enter = empty): ",
        "ask_out_base": "Output base filename (without .pdf) [{default}]: ",
        "no_cards": (
            "No cards found in folder:\n{folder}\n\n"
            "Expected filenames ending with 'a' or 'b' (e.g. card01a.png / card01b.png) "
            "OR ending with '[face,<n>]' / '[back,<n>]' (e.g. card01[face,001].png / card01[back,001].png)."
        ),
        "no_cards_examples": "Examples found in folder: {files}{more}",
        "no_cards_no_images": "No image files (.png/.jpg/.jpeg) found in the folder.",
        "done": "Done! PDF created: {path}",
        "skip_2x3": "Skipping 2x3: card images do not have bleed (min {minw}x{minh}) or are mixed.",
        "skip_gutterfold_no_backs": "Skipping Gutterfold: no backs found and no file named '{name}' in the card folder.",
        "using_cardback": "No backs found – using '{file}' as a shared card back for all cards.",
        "skip_gutterfold_missing_backs": "Skipping Gutterfold: not all fronts have a back. Missing backs for: {missing}",
        "skip_bleed_due_to_small": "Bleed will not be generated: at least one file is below the bleed minimum size {minw}x{minh} pixels ({count} file(s)). These files will be skipped: {files}",
        "warn_too_small_upscale": "Note: At least one card image is smaller than {minw}x{minh} pixels ({count} file(s)). It will be upscaled to this size: {files}",
        "error_gutterfold_space": (
            "Not enough space for Gutterfold: available {avail:.1f} pt, "
            "required {need:.1f} pt (margin: {margin} cm; top reserve: {top:.1f} pt; "
            "bottom reserve: {bottom:.1f} pt)."
        ),
        "rulebook_found": "Rulebook pages found: {files}",
        "rulebook_not_found": "No rulebook pages found (looked for '{name}*').",
        "rulebook_will_prepend": "The found rulebook pages will be prepended to the PDF.",
        "logo_found": "Logo found: {file}",
        "logo_not_found": "No logo found (looked for '{name}'). Continuing without logo.",
        "count_mismatch_warn": "Mismatched count for '{base}': face={face} back={back} -> using face={use}",
        "using_pdfconfig": "Using pdfConfig.txt from:\n{path}\nAll UI prompts will be skipped; values are taken from this file.",
        "config_title": "Configuration",
        "cfg_format": "Card format: {id} ({name})",
        "cfg_layouts": "Layout(s): {layouts}",
        "cfg_paper": "Paper: {paper}",
        "cfg_quality": "Quality: {quality}",
        "cfg_bottom_text": "Bottom text: {text}",
        "cfg_version": "Version: {version}",
        "cfg_output_name": "Output name: {name}",
        "none": "— none —",
        "exit_ok": "Done. Program finished.",
        "exit_err_header": "[ERROR] An unexpected error occurred:",
        "exit_press_enter": "Press Enter to close…",
    },
    "fr": {
        "choose_layout": "Choisissez un layout ({opts}) [All] : ",
        "startup_license": "Logiciel libre – voir LICENSE.txt (et l’avis de licence dans l'en-tête).",
        "header_welcome": "Bienvenue",
        "no_cards_title": "Aucune carte trouvée",
        "invalid_layout": "Veuillez saisir l’un des layouts proposés.",
        "format_info_note": "'Standard' utilise des images internes sans fond perdu. 'Bleed' nécessite des images avec fond perdu. 'Gutterfold' crée une mise en page pliée avec alignement recto/verso.",
        "skip_2x5": "2x5 ignoré : les images n’ont pas de fond perdu (min {minw}x{minh}) ou sont mélangées.",
        "format_info_header": "Format de carte sélectionné : {name} ({w} x {h} mm)",
        "format_info_sizes": "Tailles d'image attendues @300dpi : intérieur {iw}x{ih} px, fond perdu {bw}x{bh} px (fond perdu = 1/8\" par côté)",
        "choose_card_format": "Choisir le format des cartes (entrer un numéro, Entrée = Poker) :",
        "choose_card_format_prompt": "Sélection [1] : ",
        "invalid_card_format": "Sélection invalide. Veuillez entrer un numéro de la liste.",
        "choose_format": "Choisir le format (A4/Letter/Both) [Both] : ",
        "invalid_format": "Veuillez entrer 'A4', 'Letter' ou 'Both'.",
        "ask_folder": "Chemin du dossier contenant les images PNG/JPG : ",
        "invalid_folder": "Chemin invalide ou ce n'est pas un dossier. Réessayez.",
        "invalid_folder_path": "Chemin invalide ou ce n'est pas un dossier : {path}\nRéessayez.",
        "ask_logo": "Optionnel : chemin du logo (Entrer = auto : logo.png/jpg dans le dossier) : ",
        "logo_invalid": "Chemin du logo invalide ou fichier introuvable. Suite sans logo.",
        "ask_quality": "Choisir la qualite (Lossless/High/Medium/Low) [High] : ",
        "invalid_quality": "Qualite invalide. Entrez Lossless, High, Medium ou Low.",
        "ask_copyright": "Texte libre en bas (optionnel, 150 caractères max) : ",
        "ask_version": "Entrer la version (Entrer = vide) : ",
        "ask_out_base": "Nom de fichier de sortie (sans .pdf) [{default}] : ",
        "no_cards": (
            "Aucune carte trouvée dans le dossier :\n{folder}\n\n"
            "Noms attendus : se terminant par 'a' ou 'b' (ex. card01a.png / card01b.png) "
            "OU se terminant par '[face,<n>]' / '[back,<n>]' (ex. card01[face,001].png / card01[back,001].png)."
        ),
        "no_cards_examples": "Exemples dans le dossier : {files}{more}",
        "no_cards_no_images": "Aucun fichier image (.png/.jpg/.jpeg) trouvé dans le dossier.",
        "done": "Terminé ! PDF créé : {path}",
        "skip_2x3": "2x3 ignoré : les images n'ont pas de fond perdu (min {minw}x{minh}) ou sont mélangées.",
        "skip_gutterfold_no_backs": "Gutterfold ignoré : aucun verso trouvé et aucun fichier nommé '{name}' dans le dossier.",
        "using_cardback": "Aucun verso trouvé – utilisation de '{file}' comme verso commun pour toutes les cartes.",
        "skip_gutterfold_missing_backs": "Gutterfold ignoré : toutes les faces n'ont pas de verso. Versos manquants pour : {missing}",
        "skip_bleed_due_to_small": "Le format avec fond perdu ne sera pas généré : au moins un fichier est en dessous du minimum {minw}x{minh} px ({count} fichier(s)). Fichiers ignorés : {files}",
        "warn_too_small_upscale": "Note : au moins une image est plus petite que {minw}x{minh} pixels ({count} fichier(s)). Elle sera agrandie à cette taille : {files}",
        "error_gutterfold_space": (
            "Espace insuffisant pour le Gutterfold : {avail:.1f} pt disponibles, "
            "il en faut {need:.1f} pt (marge : {margin} cm ; réserve supérieure : {top:.1f} pt ; "
            "réserve inférieure : {bottom:.1f} pt)."
        ),
        "rulebook_found": "Pages de livret de règles trouvées : {files}",
        "rulebook_not_found": "Aucune page de livret trouvée (recherche sur '{name}*').",
        "rulebook_will_prepend": "Les pages de livret trouvées seront ajoutées au début du PDF.",
        "logo_found": "Logo trouvé : {file}",
        "logo_not_found": "Aucun logo trouvé (recherché '{name}'). Suite sans logo.",
        "count_mismatch_warn": "Quantité différente pour '{base}' : face={face} back={back} -> utilisation de face={use}",
        "using_pdfconfig": "Utilisation de pdfConfig.txt depuis :\n{path}\nToutes les invites UI seront ignorées ; les valeurs proviennent de ce fichier.",
        "config_title": "Configuration",
        "cfg_format": "Format de carte : {id} ({name})",
        "cfg_layouts": "Mise en page : {layouts}",
        "cfg_paper": "Papier : {paper}",
        "cfg_quality": "Qualité : {quality}",
        "cfg_bottom_text": "Texte bas de page : {text}",
        "cfg_version": "Version : {version}",
        "cfg_output_name": "Nom de sortie : {name}",
        "none": "— aucun —",
        "exit_ok": "Terminé. Fin du programme.",
        "exit_err_header": "[ERREUR] Une erreur inattendue s’est produite :",
        "exit_press_enter": "Appuyez sur Entrée pour fermer…",
    },
    "es": {
        "choose_layout": "Elija un layout ({opts}) [All]: ",
        "startup_license": "Software libre – véase LICENSE.txt (y el aviso de licencia del encabezado).",
        "header_welcome": "Bienvenido",
        "no_cards_title": "No se encontraron cartas",
        "invalid_layout": "Introduzca uno de los layouts ofrecidos.",
        "format_info_note": "'Standard' utiliza imágenes internas sin sangrado. 'Bleed' requiere imágenes con sangrado. 'Gutterfold' crea un diseño plegado con alineación anverso/reverso.",
        "skip_2x5": "Omitiendo 2x5: las imágenes no tienen sangrado (mín {minw}x{minh}) o están mezcladas.",
        "format_info_header": "Formato de carta seleccionado: {name} ({w} x {h} mm)",
        "format_info_sizes": "Tamaños de imagen esperados @300dpi: interior {iw}x{ih} px, sangrado {bw}x{bh} px (sangrado = 1/8\" por lado)",
        "choose_card_format": "Elegir formato de carta (número, Enter = Poker):",
        "choose_card_format_prompt": "Selección [1]: ",
        "invalid_card_format": "Selección inválida. Introduce un número de la lista.",
        "choose_format": "Elegir formato (A4/Letter/Both) [Both]: ",
        "invalid_format": "Por favor, introduce 'A4', 'Letter' o 'Both'.",
        "ask_folder": "Ruta a la carpeta con imágenes PNG/JPG: ",
        "invalid_folder": "Ruta inválida o no es una carpeta. Inténtalo de nuevo.",
        "invalid_folder_path": "Ruta inválida o no es una carpeta: {path}\nInténtalo de nuevo.",
        "ask_logo": "Opcional: ruta del logo (Enter = auto: logo.png/jpg en la carpeta): ",
        "logo_invalid": "Ruta de logo inválida o archivo no encontrado. Continuando sin logo.",
        "ask_quality": "Elegir calidad (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Calidad inválida. Introduce Lossless, High, Medium o Low.",
        "ask_copyright": "Texto libre abajo (opcional, máx. 150 caracteres): ",
        "ask_version": "Introduce versión (Enter = vacío): ",
        "ask_out_base": "Nombre base de salida (sin .pdf) [{default}]: ",
        "no_cards": (
            "No se encontraron cartas en la carpeta:\n{folder}\n\n"
            "Se esperan nombres que terminen en 'a' o 'b' (p. ej. card01a.png / card01b.png) "
            "O que terminen en '[face,<n>]' / '[back,<n>]' (p. ej. card01[face,001].png / card01[back,001].png)."
        ),
        "no_cards_examples": "Ejemplos en la carpeta: {files}{more}",
        "no_cards_no_images": "No se encontraron archivos de imagen (.png/.jpg/.jpeg) en la carpeta.",
        "done": "¡Listo! PDF creado: {path}",
        "skip_2x3": "Omitiendo 2x3: las imágenes no tienen sangrado (mín {minw}x{minh}) o están mezcladas.",
        "skip_gutterfold_no_backs": "Se omite Gutterfold: no se encontraron reversos y no hay un archivo llamado '{name}' en la carpeta.",
        "using_cardback": "No se encontraron reversos: se usa '{file}' como reverso común para todas las cartas.",
        "skip_gutterfold_missing_backs": "Se omite Gutterfold: no todas las caras tienen reverso. Reversos faltantes para: {missing}",
        "skip_bleed_due_to_small": "No se generará el formato con sangrado: al menos un archivo está por debajo del mínimo {minw}x{minh} píxeles ({count} archivo(s)). Archivos omitidos: {files}",
        "warn_too_small_upscale": "Nota: al menos una imagen es más pequeña que {minw}x{minh} píxeles ({count} archivo(s)). Se ampliará a ese tamaño: {files}",
        "error_gutterfold_space": (
            "No hay suficiente espacio para el Gutterfold: {avail:.1f} pt disponibles, "
            "se requieren {need:.1f} pt (margen: {margin} cm; reserva superior: {top:.1f} pt; "
            "reserva inferior: {bottom:.1f} pt)."
        ),
        "rulebook_found": "Páginas del reglamento encontradas: {files}",
        "rulebook_not_found": "No se encontraron páginas del reglamento (buscado '{name}*').",
        "rulebook_will_prepend": "Las páginas del reglamento se añadirán al inicio del PDF.",
        "logo_found": "Logo encontrado: {file}",
        "logo_not_found": "No se encontró logo (buscado '{name}'). Se continúa sin logo.",
        "count_mismatch_warn": "Cantidad distinta para '{base}': face={face} back={back} -> se usa face={use}",
        "using_pdfconfig": "Usando pdfConfig.txt desde:\n{path}\nTodos los diálogos de la interfaz se omitirán; los valores provienen de este archivo.",
        "config_title": "Configuración",
        "cfg_format": "Formato de carta: {id} ({name})",
        "cfg_layouts": "Diseño(s): {layouts}",
        "cfg_paper": "Papel: {paper}",
        "cfg_quality": "Calidad: {quality}",
        "cfg_bottom_text": "Texto inferior: {text}",
        "cfg_version": "Versión: {version}",
        "cfg_output_name": "Nombre de salida: {name}",
        "none": "— ninguno —",
        "exit_ok": "Listo. Fin del programa.",
        "exit_err_header": "[ERROR] Se produjo un error inesperado:",
        "exit_press_enter": "Pulsa Enter para cerrar…",
    },
    "it": {
        "choose_layout": "Scegli un layout ({opts}) [All]: ",
        "startup_license": "Software libero – vedere LICENSE.txt (e l’avviso di licenza nell'intestazione).",
        "header_welcome": "Benvenuto",
        "no_cards_title": "Nessuna carta trovata",
        "invalid_layout": "Inserire uno dei layout proposti.",
        "format_info_note": "'Standard' utilizza immagini interne senza abbondanza. 'Bleed' richiede immagini con abbondanza. 'Gutterfold' crea un layout piegato con allineamento fronte/retro.",
        "skip_2x5": "Salto 2x5: le immagini non hanno abbondanza (min {minw}x{minh}) o sono miste.",
        "format_info_header": "Formato carta selezionato: {name} ({w} x {h} mm)",
        "format_info_sizes": "Dimensioni immagine attese @300dpi: interno {iw}x{ih} px, abbondanza {bw}x{bh} px (abbondanza = 1/8\" per lato)",
        "choose_card_format": "Scegli il formato delle carte (numero, Invio = Poker):",
        "choose_card_format_prompt": "Selezione [1]: ",
        "invalid_card_format": "Selezione non valida. Inserisci un numero dalla lista.",
        "choose_format": "Scegli formato (A4/Letter/Both) [Both]: ",
        "invalid_format": "Inserisci 'A4', 'Letter' o 'Both'.",
        "ask_folder": "Percorso della cartella con immagini PNG/JPG: ",
        "invalid_folder": "Percorso non valido o non è una cartella. Riprova.",
        "invalid_folder_path": "Percorso non valido o non è una cartella: {path}\nRiprova.",
        "ask_logo": "Opzionale: percorso logo (Invio = auto: logo.png/jpg nella cartella): ",
        "logo_invalid": "Percorso logo non valido o file non trovato. Continuo senza logo.",
        "ask_quality": "Scegli qualita (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Qualita non valida. Inserisci Lossless, High, Medium o Low.",
        "ask_copyright": "Testo libero in basso (opzionale, max 150 caratteri): ",
        "ask_version": "Inserisci versione (Invio = vuoto): ",
        "ask_out_base": "Nome base output (senza .pdf) [{default}]: ",
        "no_cards": (
            "Nessuna carta trovata nella cartella:\n{folder}\n\n"
            "Nomi attesi: terminano con 'a' o 'b' (es. card01a.png / card01b.png) "
            "OPPURE terminano con '[face,<n>]' / '[back,<n>]' (es. card01[face,001].png / card01[back,001].png)."
        ),
        "no_cards_examples": "Esempi nella cartella: {files}{more}",
        "no_cards_no_images": "Nessun file immagine (.png/.jpg/.jpeg) trovato nella cartella.",
        "done": "Fatto! PDF creato: {path}",
        "skip_2x3": "Salto 2x3: le immagini non hanno abbondanza (min {minw}x{minh}) o sono miste.",
        "skip_gutterfold_no_backs": "Salto Gutterfold: nessun retro trovato e nessun file chiamato '{name}' nella cartella.",
        "using_cardback": "Nessun retro trovato – uso '{file}' come retro comune per tutte le carte.",
        "skip_gutterfold_missing_backs": "Salto Gutterfold: non tutte le fronti hanno un retro. Retro mancanti per: {missing}",
        "skip_bleed_due_to_small": "Il layout con abbondanza non verrà generato: almeno un file è sotto il minimo {minw}x{minh} pixel ({count} file). File saltati: {files}",
        "warn_too_small_upscale": "Nota: almeno un'immagine è più piccola di {minw}x{minh} pixel ({count} file). Verrà ingrandita a questa dimensione: {files}",
        "error_gutterfold_space": (
            "Spazio insufficiente per il Gutterfold: disponibili {avail:.1f} pt, "
            "necessari {need:.1f} pt (margine: {margin} cm; riserva superiore: {top:.1f} pt; "
            "riserva inferiore: {bottom:.1f} pt)."
        ),
        "rulebook_found": "Pagine del rulebook trovate: {files}",
        "rulebook_not_found": "Nessuna pagina del rulebook trovata (cercato '{name}*').",
        "rulebook_will_prepend": "Le pagine del rulebook verranno inserite all'inizio del PDF.",
        "logo_found": "Logo trovato: {file}",
        "logo_not_found": "Nessun logo trovato (cercato '{name}'). Si procede senza logo.",
        "count_mismatch_warn": "Quantità diversa per '{base}': face={face} back={back} -> uso face={use}",
        "using_pdfconfig": "Utilizzo pdfConfig.txt da:\n{path}\nTutte le richieste UI saranno ignorate; i valori provengono da questo file.",
        "config_title": "Configurazione",
        "cfg_format": "Formato carta: {id} ({name})",
        "cfg_layouts": "Layout: {layouts}",
        "cfg_paper": "Carta: {paper}",
        "cfg_quality": "Qualità: {quality}",
        "cfg_bottom_text": "Testo a piè di pagina: {text}",
        "cfg_version": "Versione: {version}",
        "cfg_output_name": "Nome output: {name}",
        "none": "— nessuno —",
        "exit_ok": "Fatto. Fine del programma.",
        "exit_err_header": "[ERRORE] Si è verificato un errore imprevisto:",
        "exit_press_enter": "Premi Invio per chiudere…",
    },
}

# =========================================================
# pdfConfig.txt  (Auto-Vorlage + Parser)
# =========================================================
PDF_CONFIG_NAME_DEFAULT = "pdfConfig.txt"

def write_pdf_config_template(dst: Path) -> None:
    """
    Write an English-only template for pdfConfig.txt.
    Created only if (a) INI does not exist and (b) the file does not yet exist.
    Field order matches the UI order exactly.
    """
    if dst.exists():
        return
    lines = [
        "# ------------------------------------------------------------",
        "# pdfConfig.txt — Template (EN only)",
        "# Copy this file into a card-image folder and adjust the values.",
        "# If present, the UI prompts are skipped and values from this file",
        "# are used for PDF generation.",
        "# ------------------------------------------------------------",
        "",
        "# 1) CARD_FORMAT (numeric id):",
        "#    1=Poker, 2=Euro, 3=Mini Euro, 4=American, 5=Mini American, 6=Custom (from INI).",
        "#    Use the numeric id listed above. If invalid/empty, 1 (Poker) is used.",
        "CARD_FORMAT=1",
        "",
        "# 2) LAYOUT:",
        "#    Allowed values (case-insensitive): Standard | Bleed | Gutterfold | All",
        "#    All = generates all supported layouts your images qualify for.",
        "LAYOUT=All",
        "",
        "# 3) PAPER:",
        "#    Allowed values (case-insensitive): Both | A4 | Letter",
        "#    Both = generate A4 and Letter variants.",
        "PAPER=Both",
        "",
        "# 4) QUALITY:",
        "#    Allowed values (case-insensitive): Lossless | High | Medium | Low",
        "#    Recommended default is High.",
        "QUALITY=High",
        "",
        "# 5) BOTTOM_TEXT:",
        "#    Optional free text printed centered in the footer (max 150 chars).",
        "#    The sequence (C) is automatically converted to ©.",
        "BOTTOM_TEXT=",
        "",
        "# 6) VERSION:",
        "#    Optional version string printed bottom-left (e.g., v1.0 or date).",
        "VERSION=",
        "",
        "# 7) OUTPUT_NAME:",
        "#    Output base filename without .pdf (invalid characters are sanitized).",
        "OUTPUT_NAME=cards",
        "",
    ]
    try:
        dst.write_text("\n".join(lines), encoding="utf-8")
    except Exception:
        pass

def read_pdf_config(path: Path) -> Dict[str, str]:
    """
    Liest eine einfache KEY=VALUE Textdatei, ignoriert leere Zeilen und #–Kommentare.
    Keys werden zu UPPERCASE normalisiert.
    """
    out: Dict[str, str] = {}
    try:
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            k, v = line.split("=", 1)
            out[k.strip().upper()] = v.strip()
    except Exception:
        return {}
    return out

def _map_layout_value(v: str) -> List[str]:
    s = (v or "").strip().lower()
    if s in ("", "all", "a"):
        return ["standard", "bleed", "gutterfold"]
    if s in ("standard", "s", "3x3", "3x4", "3"):
        return ["standard"]
    if s in ("bleed", "b", "2x3", "2x5", "2"):
        return ["bleed"]
    if s in ("gutterfold", "g", "gf"):
        return ["gutterfold"]
    # Fallback
    return ["standard", "bleed", "gutterfold"]

def _map_quality_value(v: str) -> str:
    s = (v or "").strip().lower()
    if s in ("lossless", "l", "loss", "0"):
        return "lossless"
    if s in ("high", "h", "1", ""):
        return "high"
    if s in ("medium", "m", "med", "2"):
        return "medium"
    if s in ("low", "lo", "3"):
        return "low"
    return "high"

def _map_paper_value(v: str):
    s = (v or "").strip().lower()
    if s in ("a4", "a"):
        return [(A4, "_A4")]
    if s in ("letter", "l"):
        return [(letter, "_Letter")]
    # default "both"
    return [(A4, "_A4"), (letter, "_Letter")]

def t(key: str, **kwargs) -> str:
    lang = I18N.get(LANG, I18N["de"])
    msg = lang.get(key, I18N["de"].get(key, key))
    return msg.format(**kwargs)

# =========================================================
# Rich-safe helpers (escape markup so literal [...] stays visible)
# =========================================================
def _rich_escape(text: str) -> str:
    """Escape Rich markup so strings containing [...] render literally in Panels."""
    try:
        from rich.markup import escape
        return escape(text)
    except Exception:
        return text

def _show_panel(message: str, title: str = "", border_style: str = "red") -> None:
    """Print a message either as Rich Panel (markup-safe) or plain text fallback."""
    try:
        if rprint and Panel:
            rprint(Panel(_rich_escape(message), title=title, border_style=border_style))
        else:
            print(message)
    except Exception:
        print(message)

def _sample_images_in_folder(folder: Path, limit: int = 5):
    """Return (shown_str, more_str, count_total) for sample supported images in folder."""
    try:
        imgs = [p.name for p in folder.iterdir()
            if p.is_file() and p.suffix.lower() in SUPPORTED_EXT]
        imgs.sort(key=str.lower)
        if not imgs:
            return "", "", 0
        shown = ", ".join(imgs[:limit])
        more = f" … (+{len(imgs)-limit})" if len(imgs) > limit else ""
        return shown, more, len(imgs)
    except Exception:
        return "", "", 0

def _build_no_cards_message(folder: Path) -> str:
    """Build localized 'no cards found' message including folder path and examples."""
    base = t("no_cards", folder=str(folder))
    shown, more, total = _sample_images_in_folder(folder, limit=5)
    if total > 0:
        base += "\n\n" + t("no_cards_examples", files=shown, more=more)
    else:
        base += "\n\n" + t("no_cards_no_images")
    return base

# =========================================================
# Console pause helper (useful for PyInstaller EXE)
# =========================================================
def pause_before_exit(message: str = "", print_message: bool = True) -> None:
    global _PAUSE_ALREADY_SHOWN
    try:
        if message and print_message:
            print(message)
            _PAUSE_ALREADY_SHOWN = True
            input('\n[Enter]')
    except Exception:
        pass

# =========================================================
# PDF canvas helper (sets document metadata)
# =========================================================
def create_pdf_canvas(out_path: Path, pagesize_tuple, author: str = ''):
    """Create ReportLab canvas and set PDF metadata."""
    c = canvas.Canvas(str(out_path), pagesize=pagesize_tuple)
    # PDF document property: Creator
    c.setCreator('Created by PnP PDF Creator')
    # PDF document property: Author (empty string if not provided)
    c.setAuthor(author or '')
    # PDF document property: Title = output filename (without extension)
    try:
        c.setTitle(Path(str(out_path)).stem)
    except Exception:
        c.setTitle('')
    # PDF document property: Subject/Description -> empty string
    c.setSubject('')
    return c

# =========================================================
# Card image preprocessing cache (lossy modes only)
# =========================================================
TMP_DIR = Path(tempfile.gettempdir()) / "card_pdf_cache"
TMP_DIR.mkdir(parents=True, exist_ok=True)

# Always clear preprocessing cache on each run (prevents stale resized images)
def clear_tmp_cache():
    try:
        for p in TMP_DIR.iterdir():
            if p.is_file():
                p.unlink(missing_ok=True)
    except Exception:
        # If cache cannot be cleared, continue gracefully
        pass

_CONVERT_CACHE: Dict[Tuple[str, str, str, str], Path] = {}

def get_image_px_size(img_path: Path) -> Optional[Tuple[int, int]]:
    if Image is None:
        return None
    try:
        with Image.open(img_path) as im:
            return im.size
    except Exception:
        return None

def target_pixels_for_box_inches(w_in: float, h_in: float, dpi: int) -> Tuple[int, int]:
    return int(round(w_in * dpi)), int(round(h_in * dpi))

def preprocess_card_image_for_pdf(img_path: Path, quality_key: str, box_inches: Tuple[float, float], crop_bleed: bool = True) -> Path:
    """ 
    Preprocess a card image for embedding into PDF.

    Rules:
    - crop_bleed=True  -> 3x3 + Gutterfold: ALWAYS end at INNER size 750x1050 (center-crop if needed)
    - crop_bleed=False -> 2x3: keep BLEED canvas 825x1125 (center-crop if larger), ratio-fix only if needed

    - lossless: save PNG
    - high/medium/low: downsample to target pixel box (based on dpi) and save JPEG
    """
    preset = QUALITY_PRESETS.get(quality_key, QUALITY_PRESETS["high"])
    dpi = preset["dpi"]
    jpeg_q = preset["jpeg_quality"]
    w_in, h_in = box_inches

    cache_key = (str(img_path.resolve()), quality_key, f"{w_in}x{h_in}", 'crop' if crop_bleed else 'nocrop')
    cached = _CONVERT_CACHE.get(cache_key)
    if cached and cached.exists():
        return cached

    # If PIL isn't available, just pass through (no cropping/resizing possible).
    if Image is None:
        _CONVERT_CACHE[cache_key] = img_path
        return img_path

    h = hashlib.md5((str(img_path.resolve()) + "\n" + quality_key + f"\n{w_in}x{h_in}").encode("utf-8")).hexdigest()
    ext = ".png" if quality_key == "lossless" else ".jpg"
    out_file = TMP_DIR / f"{img_path.stem}_{quality_key}_{h}{ext}"
    if out_file.exists():
        _CONVERT_CACHE[cache_key] = out_file
        return out_file

    def _center_crop_exact(im_, tw: int, th: int):
        """Center-crop to exact (tw x th). Requires im_ to be at least that large."""
        if im_.width == tw and im_.height == th:
            return im_
        left = (im_.width - tw) // 2
        top = (im_.height - th) // 2
        return im_.crop((left, top, left + tw, top + th))

    def _dbg(msg: str):
        if DEBUG_PREPROCESS:
            print(msg)

    try:
        with Image.open(img_path) as im:
            _dbg(f"[DEBUG] {img_path.name}: opened {im.width}x{im.height}, mode={im.mode}, crop_bleed={crop_bleed}, quality={quality_key}, dpi={dpi}")

            # transparency -> white background
            if im.mode in ("RGBA", "LA") or ("transparency" in im.info):
                base = Image.new("RGB", im.size, (255, 255, 255))
                im_rgba = im.convert("RGBA")
                base.paste(im_rgba, mask=im_rgba.split()[-1])
                im = base
            else:
                im = im.convert("RGB")

            if crop_bleed:
                # Target INNER (750x1050). NEVER aspect-crop to bleed ratio here.

                if im.width == BLEED_W_PX and im.height == BLEED_H_PX:
                    # exact bleed canvas -> remove fixed borders
                    im = im.crop((BLEED_LEFT_TOP_PX, BLEED_LEFT_TOP_PX,
                                  im.width - BLEED_RIGHT_BOTTOM_PX,
                                  im.height - BLEED_RIGHT_BOTTOM_PX))
                    _dbg(f"[DEBUG]   after fixed-bleed-crop: {im.width}x{im.height}")

                elif im.width >= BLEED_W_PX and im.height >= BLEED_H_PX:
                    # larger-than-bleed exports -> proportional border crop, then enforce INNER
                    left = int(round(im.width * (BLEED_LEFT_TOP_PX / BLEED_W_PX)))
                    top = int(round(im.height * (BLEED_LEFT_TOP_PX / BLEED_H_PX)))
                    right = im.width - int(round(im.width * (BLEED_RIGHT_BOTTOM_PX / BLEED_W_PX)))
                    bottom = im.height - int(round(im.height * (BLEED_RIGHT_BOTTOM_PX / BLEED_H_PX)))
                    im = im.crop((left, top, right, bottom))
                    _dbg(f"[DEBUG]   after proportional-bleed-crop: {im.width}x{im.height}")

                # If we're still larger than INNER, center-crop to exact INNER.
                if im.width >= INNER_W_PX and im.height >= INNER_H_PX and (im.width != INNER_W_PX or im.height != INNER_H_PX):
                    im = _center_crop_exact(im, INNER_W_PX, INNER_H_PX)
                    _dbg(f"[DEBUG]   after inner-enforce: {im.width}x{im.height}")

                # If image is already exactly INNER, it stays unchanged.
                # NEW: If image is smaller than INNER, upscale (stretch) to exact INNER size.
                # This avoids aborting on small images and ensures consistent placement.
                if im.width < INNER_W_PX or im.height < INNER_H_PX:
                    im = im.resize((INNER_W_PX, INNER_H_PX), resample=Image.LANCZOS)
                    _dbg(f"[DEBUG] after upscaling to INNER: {im.width}x{im.height}")


            else:
                # Target BLEED (825x1125). Keep bleed; ratio-fix only if necessary.

                if im.width >= BLEED_W_PX and im.height >= BLEED_H_PX and (im.width != BLEED_W_PX or im.height != BLEED_H_PX):
                    im = _center_crop_exact(im, BLEED_W_PX, BLEED_H_PX)
                    _dbg(f"[DEBUG]   after bleed-enforce: {im.width}x{im.height}")

                # If aspect ratio is off, center-crop to the bleed aspect ratio (11:15).
                if im.width * BLEED_H_PX != im.height * BLEED_W_PX:
                    target_ratio = BLEED_W_PX / BLEED_H_PX
                    current_ratio = im.width / im.height if im.height else target_ratio
                    if current_ratio > target_ratio:
                        new_w = int(round(im.height * target_ratio))
                        left = (im.width - new_w) // 2
                        im = im.crop((left, 0, left + new_w, im.height))
                    else:
                        new_h = int(round(im.width / target_ratio))
                        top = (im.height - new_h) // 2
                        im = im.crop((0, top, im.width, top + new_h))
                    _dbg(f"[DEBUG]   after ratio-fix (bleed): {im.width}x{im.height}")

            if quality_key == "lossless":
                im.save(out_file, "PNG", optimize=True)
                _CONVERT_CACHE[cache_key] = out_file
                _dbg(f"[DEBUG]   saved lossless: {out_file.name} -> {im.width}x{im.height}")
                return out_file

            target_w, target_h = target_pixels_for_box_inches(w_in, h_in, dpi)
            _dbg(f"[DEBUG]   target pixels: {target_w}x{target_h}")
            if im.width > target_w or im.height > target_h:
                im.thumbnail((target_w, target_h), resample=Image.LANCZOS)
                _dbg(f"[DEBUG]   after thumbnail: {im.width}x{im.height}")
            im.save(out_file, "JPEG", quality=jpeg_q, optimize=True)
            _dbg(f"[DEBUG]   saved jpeg: {out_file.name} -> {im.width}x{im.height}")

    except Exception as e:
        _CONVERT_CACHE[cache_key] = img_path
        _dbg(f"[DEBUG]   ERROR preprocessing {img_path.name}: {e}")
        return img_path

    _CONVERT_CACHE[cache_key] = out_file
    return out_file

# =========================================================
# NEU: Teil-Bleed nur an ausgewählten Außenkanten stehen lassen
# =========================================================
def preprocess_card_image_outer_bleed(
    img_path: Path,
    quality_key: str,
    keep_left_px: int,
    keep_right_px: int,
    keep_top_px: int,
    keep_bottom_px: int,
    rotate_degrees: int = 0
) -> Path:
    """
    Erzeugt ein Bild, dessen Innenfläche exakt INNER_W/H_PX bleibt, aber an
    den angegebenen Außenkanten (links/rechts/oben/unten) einen dünnen Bleed
    (z. B. 10 px) stehen lässt. Nur wenn das Quellbild echtes Bleed (>= 825x1125)
    hat; sonst fällt die Funktion automatisch auf Innenmaß-only zurück.
    Optionales rotieren (0/180) z. B. für Gutterfold-Rückseiten.
    """
    preset = QUALITY_PRESETS.get(quality_key, QUALITY_PRESETS["high"])
    jpeg_q = preset["jpeg_quality"]

    cache_key = (
        str(img_path.resolve()),
        quality_key,
        f"outerbleed:{keep_left_px}-{keep_right_px}-{keep_top_px}-{keep_bottom_px}",
        f"rot{rotate_degrees}"
    )
    cached = _CONVERT_CACHE.get(cache_key)
    if cached and cached.exists():
        return cached

    if Image is None:
        _CONVERT_CACHE[cache_key] = img_path
        return img_path

    try:
        with Image.open(img_path) as im:
            # Transparenz -> Weiß
            if im.mode in ("RGBA", "LA") or ("transparency" in im.info):
                base = Image.new("RGB", im.size, (255, 255, 255))
                im_rgba = im.convert("RGBA")
                base.paste(im_rgba, mask=im_rgba.split()[-1])
                im = base
            else:
                im = im.convert("RGB")

            # Optional vorverarbeiten: 0/180 Grad
            if rotate_degrees % 360 != 0:
                im = im.rotate(rotate_degrees % 360, expand=True)

            has_bleed = (im.width >= BLEED_W_PX and im.height >= BLEED_H_PX)

            if has_bleed:
                # Auf exakte Bleed-Canvas zentriert bringen
                if im.width != BLEED_W_PX or im.height != BLEED_H_PX:
                    left = (im.width - BLEED_W_PX) // 2
                    top  = (im.height - BLEED_H_PX) // 2
                    im = im.crop((left, top, left + BLEED_W_PX, top + BLEED_H_PX))

                # Standardmäßig würdest du 37/38 px abschneiden -> wir lassen an Außenkanten etwas stehen.
                l_cut = max(0, BLEED_LEFT_TOP_PX   - min(keep_left_px,   BLEED_LEFT_TOP_PX))
                t_cut = max(0, BLEED_LEFT_TOP_PX   - min(keep_top_px,    BLEED_LEFT_TOP_PX))
                r_cut = max(0, BLEED_RIGHT_BOTTOM_PX - min(keep_right_px,  BLEED_RIGHT_BOTTOM_PX))
                b_cut = max(0, BLEED_RIGHT_BOTTOM_PX - min(keep_bottom_px, BLEED_RIGHT_BOTTOM_PX))

                im = im.crop((l_cut, t_cut, im.width - r_cut, im.height - b_cut))

                # Zielgröße inkl. stehen gelassenem Bleed
                target_w = INNER_W_PX + min(keep_left_px, BLEED_LEFT_TOP_PX) + min(keep_right_px, BLEED_RIGHT_BOTTOM_PX)
                target_h = INNER_H_PX + min(keep_top_px,  BLEED_LEFT_TOP_PX) + min(keep_bottom_px, BLEED_RIGHT_BOTTOM_PX)

                # Falls größer -> zentriert auf Ziel beschneiden
                if im.width > target_w or im.height > target_h:
                    cx = (im.width  - target_w) // 2
                    cy = (im.height - target_h) // 2
                    im = im.crop((cx, cy, cx + target_w, cy + target_h))

                # Falls kleiner/abweichend -> exakt auf Ziel skalieren
                if im.width != target_w or im.height != target_h:
                    im = im.resize((target_w, target_h), resample=Image.LANCZOS)
            else:
                # Kein Bleed -> Innenmaß erzwingen
                if im.width >= INNER_W_PX and im.height >= INNER_H_PX:
                    cx = (im.width  - INNER_W_PX) // 2
                    cy = (im.height - INNER_H_PX) // 2
                    im = im.crop((cx, cy, cx + INNER_W_PX, cy + INNER_H_PX))
                else:
                    im = im.resize((INNER_W_PX, INNER_H_PX), resample=Image.LANCZOS)

            # Ausgabe (lossless PNG, sonst JPEG)
            h = hashlib.md5("".join(map(str, cache_key)).encode("utf-8")).hexdigest()
            ext = ".png" if quality_key == "lossless" else ".jpg"
            out_file = TMP_DIR / f"{img_path.stem}_outerbleed_{h}{ext}"
            if quality_key == "lossless":
                im.save(out_file, "PNG", optimize=True)
            else:
                im.save(out_file, "JPEG", quality=jpeg_q, optimize=True)

            _CONVERT_CACHE[cache_key] = out_file
            return out_file
    except Exception:
        _CONVERT_CACHE[cache_key] = img_path
        return img_path

# Zeichnen mit exakter Innen-Mapping-Skalierung, Bleed steht außen
def draw_card_outer_bleed(
    c: canvas.Canvas,
    processed_path: Path,
    x: float, y: float,
    card_w: float, card_h: float,
    keep_left_px: int, keep_right_px: int,
    keep_top_px: int, keep_bottom_px: int
):
    # Innenfläche muss exakt die Kartenbox füllen – so bleibt Außen-Bleed sichtbar.
    # Kleinste Rundungsunterschiede zwischen px- und pt-Geometrie dürfen
    # den Außen-Bleed nicht "auffressen".
    s_w = card_w / float(INNER_W_PX)
    s_h = card_h / float(INNER_H_PX)
    # Priorisiere Breite; wenn die Höhenabweichung spürbar wird, nimm Höhe:
    s = s_w
    if abs((s * INNER_H_PX) - card_h) > 0.5:  # Toleranz ~0,5 pt
        s = s_h
    total_w = s * (INNER_W_PX + keep_left_px + keep_right_px)
    total_h = s * (INNER_H_PX + keep_top_px + keep_bottom_px)

    # Außen-Bleed ragt aus dem Grid heraus
    dx = x - s * keep_left_px
    dy = y - s * keep_bottom_px


    c.drawImage(
        ImageReader(str(processed_path)),
        dx, dy,
        width=total_w, height=total_h,
        preserveAspectRatio=True, mask="auto"
    )

# =========================================================
# Layout capability checks (mixed image sizes)
# =========================================================
def analyze_card_images(pairs: List[Tuple[str, Optional[Path], Optional[Path]]]):
    """
    Inspect card images (only the a/b card files, NOT logo) and return:
    - sizes: dict Path->(w,h)
    - too_small: list of Paths where w<INNER_W_PX or h<INNER_H_PX (inner/trim threshold)
    - too_small_bleed: list of Paths where w<BLEED_W_PX or h<BLEED_H_PX (bleed threshold)
    - eligible_2x3_pairs: list of pairs where all existing sides are >= BLEED_W_PX x BLEED_H_PX
    - skipped_2x3_count: how many pairs were excluded from 2x3
    If PIL is not available, returns (None, [], [], pairs, 0).
    """

    if Image is None:
        # Cannot inspect sizes without PIL -> safest behavior: disable bleed and skip size-based cropping checks.
        return None, [], [], [], len(pairs)
    sizes = {}
    too_small = []
    too_small_bleed = []
    # cache sizes
    def get_size(p: Path):
        if p in sizes:
            return sizes[p]
        try:
            with Image.open(p) as im:
                sizes[p] = im.size
        except Exception:
            sizes[p] = (0, 0)
        return sizes[p]

    # check minimum size and build bleed eligible list
    eligible = []
    skipped = 0
    for base, a, b in pairs:
        # minimum check
        for p in (a, b):
            if p is None or not p.exists():
                continue
            w, h = get_size(p)
            if w < INNER_W_PX or h < INNER_H_PX:
                too_small.append(p)
                
            if w < BLEED_W_PX or h < BLEED_H_PX:
                too_small_bleed.append(p)
        # 2x3 eligibility: all existing sides must have bleed dimensions
        ok_bleed = True
        for p in (a, b):
            if p is None or not p.exists():
                continue
            w, h = get_size(p)
            if w < BLEED_W_PX or h < BLEED_H_PX:
                ok_bleed = False
                break
        if ok_bleed:
            eligible.append((base, a, b))
        else:
            skipped += 1
    return sizes, too_small, too_small_bleed, eligible, skipped

# =========================================================
# Prompts
# =========================================================

def _q_select(title: str, choices, default):
    """
    Robuster Wrapper für questionary.select(...).ask():
    - fängt Exceptions (TTY/Windows/Abbruch) ab
    - liefert bei None/leerem String immer den Default zurück
    - 'choices' und 'default' können Strings oder questionary.Choice sein
    """
    if questionary is None:
        return default
    try:
        picked = questionary.select(title, choices=choices, default=default).ask()
    except Exception:
        picked = None
    return picked if picked else default

def prompt_layout_dynamic(args=None) -> List[str]:
    # 1) CLI-Override
    if args and getattr(args, "layout", None):
        raw = args.layout.strip().lower()
        if raw in ("", "all", "a"):
            return ["standard", "bleed", "gutterfold"]
        if raw in ("standard", "s", "3x3", "3x4", "3"):
            return ["standard"]
        if raw in ("bleed", "b", "2x3", "2x5", "2"):
            return ["bleed"]
        if raw in ("gutterfold", "g", "gf"):
            return ["gutterfold"]
        print(t("invalid_layout"))
    # 2) Komfort: List-Prompt (falls questionary vorhanden)
    if questionary is not None:
        # Titel lokalisiert; Choices bleiben sprachneutral, da die Logik auf diese Keys mappt
        q_title = t("choose_layout", opts="Standard/Bleed/Gutterfold/All")   
        picked = _q_select(q_title, choices=["All", "Standard", "Bleed", "Gutterfold"], default="All")
        mapping = {
            "all": ["standard","bleed","gutterfold"],
            "standard": ["standard"],
            "bleed": ["bleed"],
            "gutterfold": ["gutterfold"],
        }
        try:
            key = picked.strip().lower() if isinstance(picked, str) else "all"
        except Exception:
            key = "all"
        return mapping.get(key, mapping["all"])        
    # 3) Fallback: bisherige Freitext-Eingabe
    opts_str = "Standard/Bleed/Gutterfold/All"
    while True:
        raw = input(t("choose_layout", opts=opts_str)).strip().lower()
        if raw in ("", "all", "a"):
            return ["standard", "bleed", "gutterfold"]
        if raw in ("standard", "s", "3x3", "3x4", "3"):
            return ["standard"]
        if raw in ("bleed", "b", "2x3", "2x5", "2"):
            return ["bleed"]
        if raw in ("gutterfold", "g", "gf"):
            return ["gutterfold"]
        print(t("invalid_layout"))

def prompt_pagesize_mode(args=None):
    # 1) CLI-Override
    if args and getattr(args, "pagesize", None):
        choice = args.pagesize.strip().lower()
        if choice in ("a4","a"):
            return [(A4, "_A4")]
        if choice in ("letter","l"):
            return [(letter, "_Letter")]
        if choice in ("both","b","a4+letter","a4letter"):
            return [(A4,"_A4"), (letter,"_Letter")]
        print(t("invalid_format"))
    # 2) Komfort: List-Prompt (falls questionary vorhanden)
    if questionary is not None:
        q_title = t("choose_format")
        picked = _q_select(q_title, choices=["Both","A4","Letter"], default="Both")
        if str(picked) == "A4":
            return [(A4, "_A4")]
        if str(picked) == "Letter":
            return [(letter, "_Letter")]
        return [(A4,"_A4"), (letter,"_Letter")]
    # 3) Fallback: bisherige Freitext-Eingabe
    while True:
        choice = input(t("choose_format")).strip().lower()
        if choice == "":
           choice = "both"
        if choice in ("a4", "a"):
            return [(A4, "_A4")]
        if choice in ("letter", "l"):
            return [(letter, "_Letter")]
        if choice in ("both", "b", "a4+letter", "a4letter"):
            return [(A4, "_A4"), (letter, "_Letter")]
        print(t("invalid_format"))

def prompt_folder() -> Path:
    while True:
        raw = input(t("ask_folder")).strip().strip('"')
        # Wichtig: Leere Eingabe niemals akzeptieren ? direkt erneut fragen
        if raw == "":
            print(t("invalid_folder_path", path=raw))
            continue

        # Erst nach nicht-leerer Eingabe den Pfad auflösen und validieren
        folder = Path(expanduser(raw)).expanduser().resolve()
        if folder.exists() and folder.is_dir():
            return folder
        print(t("invalid_folder_path", path=raw))

def prompt_logo_path(folder: Path) -> Optional[Path]:
    p = input(t("ask_logo")).strip().strip('"')
    if p:
        lp = Path(expanduser(p)).expanduser().resolve()
        if lp.exists() and lp.is_file() and lp.suffix.lower() in SUPPORTED_EXT:
            return lp
        print(t("logo_invalid"))
        return None

    for name in ("logo.png", "logo.jpg", "logo.jpeg"):
        cand = folder / name
        if cand.exists() and cand.is_file():
            return cand
    for name in ("logo.png", "logo.jpg", "logo.jpeg"):
        cand = Path.cwd() / name
        if cand.exists() and cand.is_file():
            return cand
    return None

def prompt_quality(args=None) -> str:
    # 1) CLI-Override
    if args and getattr(args, "quality", None):
        raw = args.quality.strip().lower()
        if raw in ("lossless", "l", "loss", "0"):
            return "lossless"
        if raw in ("high", "h", "1"):
            return "high"
        if raw in ("medium", "m", "med", "2"):
            return "medium"
        if raw in ("low", "lo", "3"):
            return "low"
        # Ungültige CLI-Eingabe -> weiter zum Prompt
    # 2) Komfort: questionary-Select (lokalisierter Titel)
    if questionary is not None:
        q_title = t("ask_quality")
        # Choices als Objekte erstellen und **dasselbe Objekt** als default verwenden
        choices = [
            questionary.Choice("Lossless", "lossless"),
            questionary.Choice("High",     "high"),
            questionary.Choice("Medium",   "medium"),
            questionary.Choice("Low",      "low"),
        ]
        try:
            picked = questionary.select(
                q_title,
                choices=choices,
                default=choices[1]  # "High" – exakt das Objekt aus der Liste
            ).ask()
        except Exception:
            picked = None
        # Frage abgebrochen? -> Fallback auf DEFAULT_QUALITY
        return picked or DEFAULT_QUALITY
    # 3) Fallback: Freitext-Prompt (lokalisiert)
    while True:
        raw = input(t("ask_quality")).strip().lower()
        if raw == "":
            return DEFAULT_QUALITY
        if raw in ("lossless", "l", "loss", "0"):
            return "lossless"
        if raw in ("high", "h", "1"):
            return "high"
        if raw in ("medium", "m", "med", "2"):
            return "medium"
        if raw in ("low", "lo", "3"):
            return "low"
        print(t("invalid_quality"))

def prompt_copyright_name() -> Optional[str]:
    raw = input(t("ask_copyright")).strip()
    if not raw:
        return None
    normalized = raw.replace("(C)", "©").replace("(c)", "©")
    return normalized[:COPY_MAX_CHARS]

def prompt_version() -> str:
    return input(t("ask_version")).strip()

def prompt_output_base(default_base: str) -> str:
    base = input(t("ask_out_base", default=default_base)).strip()
    return base if base else default_base


# =========================================================
# Card pairing
# =========================================================
def find_card_pairs(folder: Path) -> List[Tuple[str, Optional[Path], Optional[Path]]]:
    """
    Find and pair card front/back images – with count support.
    Erweiterungen:
    - Bracket-Schema: base[face,NNN].png / base[back,NNN].png
      -> NNN ist die gewünschte Stückzahl (1..3 Ziffern).
      -> Face/Back-Counts unterschiedlich: Warnung via I18N; Face-Count gewinnt.
      -> Nur Face vorhanden: wird 'count' mal ohne Back aufgenommen.
    - Legacy '...a'/'...b': Count = 1 (wie bisher).
    Rückgabe ist eine expandierte Liste, in der jedes Tupel eine physische Karte repräsentiert.
    """
    files = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXT]

    # Patterns
    # Legacy: ...a / ...b
    # ab_pattern = re.compile(r"^(.*?)(?:[_\-\s]?)([ab])$", re.IGNORECASE)
    ab_pattern = re.compile(r"^(.*?)(?:[\_\\-\\s]?)([ab])$", re.IGNORECASE)
    

    # New scheme: base[face,NNN] OR base[back,NNN]
    # Wichtig: base = alles VOR der ersten Klammer!
    bracket_pattern = re.compile(r"^(.*?)\[(face|back),(\d{1,3})\]$", re.IGNORECASE)

    # Map: base -> entry
    # Wichtig: NUM wird nicht mehr zum Key!
    pairs_map: Dict[str, Dict[str, object]] = {}

    for f in files:
        stem = f.stem

        m2 = bracket_pattern.match(stem)
        if m2:
            base = m2.group(1)                 # nur VOR der Klammer
            kind = m2.group(2).lower()         # face/back
            num_raw = m2.group(3)              # NNN (1–3 digits)
            count_val = max(1, min(int(num_raw), 999))

            key = base.lower()                 # KEY = NUR DER BASENAME

            entry = pairs_map.setdefault(key, {
                'base': base,
                'num': None,
                'face': None,
                'back': None,
                'face_count': None,
                'back_count': None,
            })
            if kind == 'face':
                entry['face'] = f
                entry['face_count'] = count_val
            else:
                entry['back'] = f
                entry['back_count'] = count_val
            continue

    # Legacy-Schema separat, damit "base__NNN" Keys nicht mit "base" kollidieren
    for f in files:
        m1 = ab_pattern.match(f.stem)
        if not m1:
            continue
        base = m1.group(1)
        side = m1.group(2).lower()
        key = base.lower()
        entry = pairs_map.setdefault(key, {
            'base': base,
            'num': None,
            'face': None,
            'back': None,
            'face_count': None,
            'back_count': None,
        })
        if side == 'a':
            entry['face'] = f
            entry['face_count'] = 1
        else:
            entry['back'] = f
            entry['back_count'] = 1

    # Sortierung: base (case-insensitiv)
    def sort_key(item):
        return item[1]["base"].lower()

    expanded: List[Tuple[str, Optional[Path], Optional[Path]]] = []

    for _key, d in sorted(pairs_map.items(), key=sort_key):
        base = str(d.get('base', ''))
        num = d.get('num')
        face: Optional[Path] = d.get('face')  # type: ignore
        back: Optional[Path] = d.get('back')  # type: ignore
        face_count = d.get('face_count') if isinstance(d.get('face_count'), int) else None
        back_count = d.get('back_count') if isinstance(d.get('back_count'), int) else None

        # Defaults
        if face and not face_count:
            face_count = 1
        if back and not back_count:
            back_count = 1

        # Count-Regeln
        if face and back:
            if face_count and back_count and face_count != back_count:
                use_count = face_count or 1
                # Lokalisierte Warnung
                print(t("count_mismatch_warn", base=base, face=face_count, back=back_count, use=use_count))
                count_to_use = use_count
            else:
                count_to_use = face_count or back_count or 1
        elif face and not back:
            count_to_use = face_count or 1
        elif back and not face:
            count_to_use = back_count or 1
        else:
            continue  # weder Face noch Back

        # Display name ohne NNN
        display_name = base
        count_to_use = max(1, int(count_to_use))
        for _ in range(count_to_use):
            expanded.append((display_name,
                             face if (face and face.exists()) else face,
                             back  if (back  and back.exists())  else back))
    return expanded

def find_named_image_in_folder(folder: Path, basename: str) -> Optional[Path]:
    """Find an image file in folder with stem matching basename (Windows-typical: case-insensitive).
    Accepts any supported extension (.png/.jpg/.jpeg). Returns first match in deterministic order.
    """
    if not basename:
        return None
    want = basename.strip()
    if not want:
        return None
    want_l = want.lower()
    ext_rank = {'.png': 0, '.jpg': 1, '.jpeg': 2}
    candidates = [p for p in folder.iterdir()
                  if p.is_file() and p.suffix.lower() in SUPPORTED_EXT and p.stem.lower() == want_l]
    if not candidates:
        return None
    candidates.sort(key=lambda p: (ext_rank.get(p.suffix.lower(), 9), p.name.lower()))
    return candidates[0]

def find_rulebook_images(folder: Path, basename: str) -> List[Path]:
    """
    Sucht nach Bildern (png/jpg/jpeg), deren Dateiname (stem) mit `basename` beginnt,
    z. B. rulebook01.png, rulebook02.jpg. Alphanumerische Sortierung.
    """
    if not basename:
        return []
    want = basename.strip().lower()
    files = [
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXT and p.stem.lower().startswith(want)
    ]
    def _alnum_key(p: Path):
        import re
        parts = re.split(r'(\d+)', p.name.lower())
        return [int(s) if s.isdigit() else s for s in parts]
    return sorted(files, key=_alnum_key)

def chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


# =========================================================
# Placement helpers
# =========================================================
def fit_image_into_box(img_path: Path, box_w: float, box_h: float) -> Tuple[float, float]:
    size = get_image_px_size(img_path)
    if not size:
        return box_w, box_h
    iw, ih = size
    if iw <= 0 or ih <= 0:
        return box_w, box_h
    scale = min(box_w / iw, box_h / ih)
    return iw * scale, ih * scale

def fit_logo_with_constraints(logo_path: Path, max_w: float, max_h: float) -> Tuple[float, float]:
    size = get_image_px_size(logo_path)
    if not size:
        return max_w, max_h
    w, h = float(size[0]), float(size[1])
    if w <= max_w and h <= max_h:
        return w, h
    scale = min(max_w / w, max_h / h)
    return w * scale, h * scale

def compute_grid_origin_centered(page_w: float, page_h: float, grid_w: float, grid_h: float) -> Tuple[float, float]:
    return (page_w - grid_w) / 2.0, (page_h - grid_h) / 2.0

# ---------------------------------------------------------
# Rulebook-Seiten (Frontmatter) einfügen
# ---------------------------------------------------------
def draw_rulebook_pages(c: canvas.Canvas, pagesize_tuple, image_paths: List[Path], mode: str = "auto", force_mode: str = "auto") -> None:
    """
    Fügt die angegebenen Bildseiten VORNE ein:
    - jeweils ganze Seite A4/Letter (entsprechend pagesize_tuple)
    - Bild zentriert
    - nur herunterskalieren, NIE hochskalieren
    - KEIN Logo, KEINE Fußzeile/Version/Seitenzahl
    - mode: "portrait_pref" (Standard), "landscape_pref" (Bleed/Gutterfold), "auto"
    - force_mode: "auto" (Standardverhalten), "off", "force_landscape", "force_portrait"
    """
    if not image_paths:
        return
    page_w, page_h = pagesize_tuple
    for p in image_paths:
        try:
            if not p.exists():
                continue
            # --- Per-Image Rotation abhängig vom Ziel-Layout ---
            # Regeln:
            #   - landscape_pref (Bleed/Gutterfold): rotate 90° rechts, wenn Höhe > Breite
            #   - portrait_pref  (Standard)        : rotate 90° rechts, wenn Breite > Höhe
            #   - quadratisch: keine Rotation
            rotated_reader = None
            iw = ih = None
            size = get_image_px_size(p)
            if size:
                iw, ih = float(size[0]), float(size[1])
            # Default: keine Rotation, direkter Reader
            img_reader = ImageReader(str(p))

            def _need_rotate_clockwise(iw_f: float, ih_f: float, mode_s: str) -> bool:
                if iw_f is None or ih_f is None:
                    return False
                if abs(iw_f - ih_f) < 1e-6:
                    return False  # quadratisch
                if mode_s == "landscape_pref":
                    return ih_f > iw_f
                if mode_s == "portrait_pref":
                    return iw_f > ih_f
                return False

            # force_mode greift zuerst
            fmode = (force_mode or "auto").lower()
            if fmode == "off":
                target_mode = "none"  # niemals rotieren
            elif fmode == "force_landscape":
                target_mode = "landscape_pref"
            elif fmode == "force_portrait":
                target_mode = "portrait_pref"
            else:
                # "auto" -> aus dem Layout-Mode ableiten
                target_mode = (mode or "auto").lower()
                if target_mode not in ("portrait_pref", "landscape_pref"):
                    target_mode = "portrait_pref"

            if Image is not None and iw is not None and ih is not None and target_mode != "none" and _need_rotate_clockwise(iw, ih, target_mode):
                # 90° rechts drehen (clockwise) -> PIL transpose ROTATE_270
                from PIL import Image as _PILImage
                with _PILImage.open(p) as _im:
                    _im = _im.convert("RGBA")
                    _im = _im.transpose(_PILImage.ROTATE_270)  # 90° clockwise
                    iw, ih = float(_im.width), float(_im.height)
                    rotated_reader = ImageReader(_im)
            # -------------------------------------------------------
            #   Skalierung: NICHT hochskalieren
            #   Nur verkleinern, wenn Bild > Seite (physisch @300dpi)
            # -------------------------------------------------------
            if iw is None or ih is None or iw <= 0 or ih <= 0:
                draw_w_pt = page_w
                draw_h_pt = page_h
            else:
                # Bildgröße in Punkten (1 pt = 1/72", Bild hat nativen 300 dpi)
                iw_pt = iw / 300.0 * 72.0
                ih_pt = ih / 300.0 * 72.0
    
                # Downscale-Faktor (niemals >1)
                scale = min(1.0, min(page_w / iw_pt, page_h / ih_pt))

                draw_w_pt = iw_pt * scale
                draw_h_pt = ih_pt * scale

            # Zentrierung
            dx = (page_w - draw_w_pt) / 2.0
            dy = (page_h - draw_h_pt) / 2.0

            c.drawImage(
                rotated_reader or img_reader,
                dx, dy,
                width=draw_w_pt,
                height=draw_h_pt,
                preserveAspectRatio=True,
                mask="auto"
            )

            c.showPage()
        except Exception:
            # robust weiter; fehlerhafte Datei überspringen
            continue

# =========================================================
# CLI / Rich / Komfort: Argumente, hübsche Ausgabe, Warm-Up
# =========================================================
def parse_args():
    p = argparse.ArgumentParser(description="PnP PDF Creator (polished CLI)")
    p.add_argument("--lang", choices=["de","en","fr","es","it"], help="UI-Sprache")
    p.add_argument("--format", dest="card_format", help="Kartenformatname (z.B. 'Poker', 'Euro', ...)")
    p.add_argument("--layout", choices=["standard","bleed","gutterfold","all"], help="Layoutwahl")
    p.add_argument("--pagesize", choices=["A4","Letter","Both"], help="Papierformat")
    p.add_argument("--folder", type=str, help="Ordner mit Kartenbildern")
    p.add_argument("--logo", type=str, help="Pfad zu Logo-Datei (optional)")
    p.add_argument("--quality", choices=["lossless","high","medium","low"], help="Qualität")
    p.add_argument("--copyright", type=str,
                   help="Copyright-Name (unten zentriert; leer = kein Copyright)")
    p.add_argument("--version", type=str, help="Versionsstring (unten links)")
    p.add_argument("--out", dest="out_base", type=str, help="Ausgabebasis (ohne .pdf)")
    return p.parse_args()

def _show_header():
    if rprint and Panel:
        rprint(Panel.fit(
            f"[bold white]PnP PDF Creator[/bold white] [green]{SCRIPT_VERSION}[/green]\n"
            f"PIL available: [cyan]{Image is not None}[/cyan]",
            title=t("header_welcome"), border_style="green"))
    else:
        print(f"PnP PDF Creator {SCRIPT_VERSION}\nPIL available: {Image is not None}")

def _show_format_table(fmt: dict):
    if not rprint or not Table:
        return
    # Tabellen-Titel & Spaltenüberschriften lokalisieren
    title = {
        "de": "Kartenformat",
        "en": "Card format",
        "fr": "Format de carte",
        "es": "Formato de carta",
        "it": "Formato carta",
    }.get(LANG, "Card format")

    col_name = {
        "de": "Name",
        "en": "Name",
        "fr": "Nom",
        "es": "Nombre",
        "it": "Nome",
    }.get(LANG, "Name")

    col_w = {
        "de": "Breite (mm)",
        "en": "Width (mm)",
        "fr": "Largeur (mm)",
        "es": "Ancho (mm)",
        "it": "Larghezza (mm)",
    }.get(LANG, "Width (mm)")

    col_h = {
        "de": "Höhe (mm)",
        "en": "Height (mm)",
        "fr": "Hauteur (mm)",
        "es": "Altura (mm)",
        "it": "Altezza (mm)",
    }.get(LANG, "Height (mm)")

    table = Table(title=title, show_header=True, header_style="bold magenta")
    table.add_column(col_name)
    table.add_column(col_w)
    table.add_column(col_h)
    # Für INI-Custom-Formate (fmt.get('src') == 'ini') stets 1 Nachkommastelle anzeigen.
    if fmt.get('src') == 'ini':
        w_str = _mm_str_custom(float(fmt["w_mm"]))
        h_str = _mm_str_custom(float(fmt["h_mm"]))
    else:
        w_str = _mm_str(float(fmt["w_mm"]))
        h_str = _mm_str(float(fmt["h_mm"]))
    table.add_row(fmt["name"], w_str, h_str)
    rprint(table)
    # Info-Panel OHNE "Hinweis:" Präfix
    info_title = {
        "de": "Info",
        "en": "Info",
        "fr": "Info",
        "es": "Información",
       "it": "Informazioni",
    }.get(LANG, "Info")

    if Panel:
        rprint(
            Panel.fit(
                t('format_info_note'),
                title=info_title,
                border_style="cyan"
            )
        )

def _collect_all_images_for(layout_key, pairs):
    # De-dupe aller relevanten Bildpfade für Warm-Up
    if layout_key in ("bleed","2x3","2x5"):
        imgs = [p for (_n,a,b) in pairs for p in (a,b) if p]
    else:
        imgs = [p for (_n,a,b) in pairs for p in (a,b) if p]
    seen, out = set(), []
    for p in imgs:
        rp = str(Path(p).resolve())
        if rp not in seen:
            seen.add(rp); out.append(Path(p))
    return out

def warmup_preprocessing(img_paths, quality_key, card_box_inches, crop_bleed):
    # Optionales Vorwärmen (zeigt Fortschritt); Zeichnen nutzt dann Cache
    if not img_paths:
        return
    if (rprint is None) or (Progress is None):
        # Kein rich installiert ? stilles Warm-Up
        for p in img_paths:
            preprocess_card_image_for_pdf(p, quality_key, card_box_inches, crop_bleed=crop_bleed)
        return
    with Progress(
        TextColumn("[bold cyan]{task.description}"),
        BarColumn(),
        TextColumn("{task.percentage:>3.0f}%"),
        TimeRemainingColumn(),
        transient=True
    ) as progress:
        task = progress.add_task("Bilder vorbereiten…", total=len(img_paths))
        for p in img_paths:
            preprocess_card_image_for_pdf(p, quality_key, card_box_inches, crop_bleed=crop_bleed)
            progress.advance(task)

# =========================================================
# Zentrierung mit druckfreiem Rand + Reserven (NEU)
# =========================================================
def compute_grid_origin_centered_with_margins(
    page_w: float, page_h: float,
    grid_w: float, grid_h: float,
    margins: Dict[str, float],
    top_reserved_pt: float,
    bottom_reserved_pt: float,
    left_reserved_pt: float = 0.0,
    right_reserved_pt: float = 0.0
) -> Tuple[float, float]:
    # Verfügbarer Bereich nach Abzug von Rändern UND Reserven
    avail_w = page_w - margins["left"] - margins["right"] - left_reserved_pt - right_reserved_pt
    avail_h = page_h - margins["top"] - margins["bottom"] - top_reserved_pt - bottom_reserved_pt

    if avail_w <= 0 or avail_h <= 0:
        # Fallback: direkt oberhalb/links der Reserven platzieren
        return margins["left"] + left_reserved_pt, margins["bottom"] + bottom_reserved_pt

    # WICHTIG: Ursprung des Zentrierbereichs über Reserven legen!
    x = margins["left"] + left_reserved_pt + (avail_w - grid_w) / 2.0
    y = margins["bottom"] + bottom_reserved_pt + (avail_h - grid_h) / 2.0
    return x, y

# =========================================================
# Maximiere cols/rows innerhalb der verfügbaren Fläche (NEU)
# =========================================================
def compute_max_grid_counts(
    page_w: float, page_h: float,
    box_w: float, box_h: float,
    margins: Dict[str, float],
    logo_path: Optional[Path],
    bottom_reserved_pt: float,
    extra_vertical_pt: float = 0.0,
    left_reserved_pt: float = RESERVE_LEFT_PT,
    right_reserved_pt: float = RESERVE_RIGHT_PT,
    top_fixed_reserved_pt: float = RESERVE_TOP_PT
) -> Tuple[int, int, float, float, float, float, float]:
   
    # WICHTIG: Grid-Zentrierung soll nicht durch Logo-Reserve verkleinert werden.
    # Daher fließt nur der FIXE Kopfbereich (top_fixed_reserved_pt) in die
    # Platzberechnung ein; das Logo wird später über dem Grid skaliert gezeichnet.
    top_res = top_fixed_reserved_pt

    avail_w = page_w - margins["left"] - margins["right"] - left_reserved_pt - right_reserved_pt
    avail_h = page_h - margins["top"] - margins["bottom"] - top_res - bottom_reserved_pt

    if avail_w <= 0 or avail_h <= 0:
        return (1, 1,
                margins["left"] + left_reserved_pt,
                margins["bottom"] + bottom_reserved_pt,
                box_w, box_h + extra_vertical_pt,
                margins["bottom"] + bottom_reserved_pt + box_h + extra_vertical_pt)

    rows = max(1, int((avail_h - extra_vertical_pt) // box_h))
    cols = max(1, int(avail_w // box_w))
    grid_w = cols * box_w
    grid_h = rows * box_h + extra_vertical_pt

    x0, y0 = compute_grid_origin_centered_with_margins(
        page_w, page_h, grid_w, grid_h,
        margins,
        top_res, bottom_reserved_pt,
        left_reserved_pt, right_reserved_pt
    )
    grid_top_y = y0 + grid_h
    return cols, rows, x0, y0, grid_w, grid_h, grid_top_y


# =========================================================
# Drawing: bottom + logo
# =========================================================
def draw_bottom_line(c: canvas.Canvas, page_w: float,
                     copyright_name: Optional[str],
                     version_str: str,
                     page_label: str,
                     y_override: Optional[float] = None):

    c.saveState()
    c.setFont("Helvetica", 9)
    y = BOTTOM_Y if y_override is None else y_override
    
    if version_str:
        c.drawString(LEFT_MARGIN, y, version_str)

    if copyright_name:
        text = copyright_name[:COPY_MAX_CHARS]
        c.drawCentredString(page_w / 2.0, y, text)

    c.drawRightString(page_w - RIGHT_MARGIN, y, page_label)
    c.restoreState()

# =========================================================
# Layout Gutterfold (example-like): 4 rows, 2 columns (a left / b right)
# - cards rotated 90°
# - back mirrored horizontally
# - dashed fold line in gutter
# - outside-only cut marks
# =========================================================

def _fit_image_into_box_rotated(img_path: Path, box_w: float, box_h: float, rotate_deg: int) -> Tuple[float, float]:
    """Return draw_w, draw_h after rotation so that rotated image fits into box."""
    size = get_image_px_size(img_path)
    if not size:
        return box_w, box_h
    iw, ih = size
    if iw <= 0 or ih <= 0:
        return box_w, box_h
    r = rotate_deg % 360
    if r in (90, 270):
        # rotated dims: w = ih, h = iw
        scale = min(box_w / ih, box_h / iw)
        return ih * scale, iw * scale
    scale = min(box_w / iw, box_h / ih)
    return iw * scale, ih * scale


def draw_image_transformed(
    c: canvas.Canvas,
    img_path: Path,
    x: float, y: float,
    box_w: float, box_h: float,
    rotate_deg: int = 0,
    mirror_x: bool = False,
):
    """Draw image centered in box; mirror_x is applied in page X axis (good for gutter folding)."""
    draw_w, draw_h = _fit_image_into_box_rotated(img_path, box_w, box_h, rotate_deg)
    cx = x + box_w / 2.0
    cy = y + box_h / 2.0
    c.saveState()
    c.translate(cx, cy)
    if mirror_x:
        c.scale(-1, 1)
    if rotate_deg:
        c.rotate(rotate_deg)
    c.drawImage(ImageReader(str(img_path)), -draw_w / 2.0, -draw_h / 2.0,
                width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
    c.restoreState()

def draw_gutterfold_line_horizontal(c: canvas.Canvas, x: float, y: float, w: float):
    c.saveState()
    c.setLineWidth(GF_FOLD_LINE_WIDTH)
    if GF_FOLD_LINE_DASH:
        c.setDash(GF_FOLD_LINE_DASH[0], GF_FOLD_LINE_DASH[1])
    from reportlab.lib.colors import black
    c.setStrokeColor(black)
    c.line(x, y, x + w, y)
    c.restoreState()

def draw_gutter_bridge_marks(
    c: canvas.Canvas,
    x_positions: List[float],
    y_gutter_bottom: float,
    y_gutter_top: float,
):
    """
    Draw vertical 'bridge' cut marks ONLY across the gutter area,
    i.e., from the top edge of the bottom row up to the bottom edge of the top row.
    """
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)
    for x in x_positions:
        c.line(x, y_gutter_bottom, x, y_gutter_top)
    c.restoreState()

def draw_cutmarks_gutterfold(
    c: canvas.Canvas,
    x0: float, y0: float,
    grid_w: float, grid_h: float,
    y_edges: List[float],
    x_marks: List[float],
):
    """Outside-only crop marks (similar visual style to your 2x3 outer marks)."""
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)
    L = CUTMARK_LEN_PT_STD
    x_left = x0
    x_right = x0 + grid_w
    y_bottom = y0
    y_top = y0 + grid_h

    for y in y_edges:
        c.line(x_left - L, y, x_left, y)
        c.line(x_right, y, x_right + L, y)

    for x in x_marks:
        c.line(x, y_bottom - L, x, y_bottom)
        c.line(x, y_top, x, y_top + L)

    c.restoreState()

# =========================================================
# Layout 3x3: inner crosses + outer marks between cards
# =========================================================
def draw_inner_crosses_grid(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float, cols: int, rows: int):
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)
    half = CUTMARK_LEN_PT_STD / 2.0
    xs = [x0 + j * card_w for j in range(1, cols)]
    ys = [y0 + i * card_h for i in range(1, rows)]
    for x in xs:
        for y in ys:
            c.line(x - half, y, x + half, y)
            c.line(x, y - half, x, y + half)
    c.restoreState()

def draw_outer_marks_grid(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float, cols: int, rows: int):
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)
    half = CUTMARK_LEN_PT_STD / 2.0
    grid_w = cols * card_w
    grid_h = rows * card_h
    xs = [x0 + j * card_w for j in range(1, cols)]
    ys = [y0 + i * card_h for i in range(1, rows)]
    y_bottom = y0
    y_top = y0 + grid_h
    x_left = x0
    x_right = x0 + grid_w
    for x in xs:
        c.line(x, y_bottom - half, x, y_bottom + half)
        c.line(x, y_top    - half, x, y_top    + half)
        c.line(x - half, y_bottom, x + half, y_bottom)
        c.line(x - half, y_top,    x + half, y_top)
    for y in ys:
        c.line(x_left  - half, y, x_left  + half, y)
        c.line(x_right - half, y, x_right + half, y)
        c.line(x_left,  y - half, x_left,  y + half)
        c.line(x_right, y - half, x_right, y + half)
    c.restoreState()

def draw_corner_marks_grid(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float, cols: int, rows: int):
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)
    half = CUTMARK_LEN_PT_STD / 2.0
    grid_w = cols * card_w
    grid_h = rows * card_h
    x_left = x0
    x_right = x0 + grid_w
    y_bottom = y0
    y_top = y0 + grid_h
    for (x, y) in ((x_left, y_bottom), (x_right, y_bottom), (x_left, y_top), (x_right, y_top)):
        c.line(x - half, y, x + half, y)
        c.line(x, y - half, x, y + half)
    c.restoreState()

def draw_inner_crosses_3x3(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float):
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)
    half = CUTMARK_LEN_PT_STD / 2.0
    xs = [x0 + card_w, x0 + 2 * card_w]
    ys = [y0 + card_h, y0 + 2 * card_h]
    for x in xs:
        for y in ys:
            c.line(x - half, y, x + half, y)
            c.line(x, y - half, x, y + half)
    c.restoreState()

def draw_outer_marks_3x3(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float):
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)
    half = CUTMARK_LEN_PT_STD / 2.0

    grid_w = 3 * card_w
    grid_h = 3 * card_h
    xs = [x0 + card_w, x0 + 2 * card_w]
    ys = [y0 + card_h, y0 + 2 * card_h]
    y_bottom = y0
    y_top = y0 + grid_h
    x_left = x0
    x_right = x0 + grid_w
    for x in xs:
        c.line(x, y_bottom - half, x, y_bottom + half)
        c.line(x, y_top    - half, x, y_top    + half)
        c.line(x - half, y_bottom, x + half, y_bottom)
        c.line(x - half, y_top,    x + half, y_top)
    for y in ys:
        c.line(x_left  - half, y, x_left  + half, y)
        c.line(x_right - half, y, x_right + half, y)
        c.line(x_left,  y - half, x_left,  y + half)
        c.line(x_right, y - half, x_right, y + half)
    c.restoreState()

def draw_corner_marks_3x3(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float):
    """
    Draw L-shaped corner marks at all 4 corners of the 3x3 grid.
    Each segment is centered on the grid corner, so it is half inside and half outside,
    matching the visual style of draw_outer_marks_3x3().
    """
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
    c.setStrokeColor(CUTMARK_COLOR)

    half = CUTMARK_LEN_PT_STD / 2.0
    grid_w = 3 * card_w
    grid_h = 3 * card_h

    x_left = x0
    x_right = x0 + grid_w
    y_bottom = y0
    y_top = y0 + grid_h

    # Bottom-left corner
    c.line(x_left - half, y_bottom, x_left + half, y_bottom)   # horizontal (half out / half in)
    c.line(x_left, y_bottom - half, x_left, y_bottom + half)   # vertical   (half out / half in)

    # Bottom-right corner
    c.line(x_right - half, y_bottom, x_right + half, y_bottom)
    c.line(x_right, y_bottom - half, x_right, y_bottom + half)

    # Top-left corner
    c.line(x_left - half, y_top, x_left + half, y_top)
    c.line(x_left, y_top - half, x_left, y_top + half)

    # Top-right corner
    c.line(x_right - half, y_top, x_right + half, y_top)
    c.line(x_right, y_top - half, x_right, y_top + half)

    c.restoreState()

def _compute_enclosing_edges(img_paths, cols, rows, is_back=False):
    """
    Ermittelt für ein (teilweise belegtes) Grid die umschließenden Kanten
    entlang der tatsächlich belegten Zellen:
      - min_row / max_row: erste/letzte belegte Zeile insgesamt
      - min_col_row[i] / max_col_row[i]: erste/letzte belegte Spalte in Zeile i
    Bei Rückseiten wird die Spaltenlage wie in der Zeichenschleife gespiegelt.
    """
    per_page = cols * rows
    occ = [[False] * cols for _ in range(rows)]

    for idx in range(min(len(img_paths), per_page)):
        p = img_paths[idx]
        if p is None or (hasattr(p, "exists") and not p.exists()):
            continue
        row = idx // cols
        col = idx % cols
        if is_back:
            col = (cols - 1) - col
        occ[row][col] = True

    row_has = [any(occ[i]) for i in range(rows)]
    # Top (erste belegte Zeile) und Bottom (letzte belegte Zeile)
    min_row = next((i for i, has in enumerate(row_has) if has), 0)
    max_row = next((i for i in range(rows - 1, -1, -1) if row_has[i]), rows - 1)

    # Erste/letzte belegte Spalte je Zeile
    min_col_row = [None] * rows
    max_col_row = [None] * rows
    for i in range(rows):
        if row_has[i]:
            min_col_row[i] = next((j for j in range(cols) if occ[i][j]), 0)
            max_col_row[i] = next((j for j in range(cols - 1, -1, -1) if occ[i][j]), cols - 1)

    return min_row, max_row, min_col_row, max_col_row

def place_images_grid_inner(
    c: canvas.Canvas,
    img_paths: List[Optional[Path]],
    x0: float, y0: float,
    card_w: float, card_h: float,
    cols: int, rows: int,
    is_back: bool,
    quality_key: str,
    card_box_inches: Tuple[float, float],
    outer_bleed_keep_px: int = 0
):
    """
    Standard-Layout (Innenmaß-Boxen) mit optionalem Außen-Bleed nur an den
    'logisch äußeren' Kanten des tatsächlich belegten Rasters.
    - Innenfläche (INNER_W/H_PX) wird exakt auf card_w/card_h gemappt.
    - Außen-Bleed steht nur außen (oben/unten/links/rechts), niemals zwischen Karten.
    - Fehlt Bleed in der Quelle, fällt die Darstellung automatisch auf Innenmaß zurück.
    - Marken werden immer nach der Bildschleife gezeichnet.
    """
    per_page = cols * rows

    # Einheitliche Skalierung, sodass die Innenfläche exakt die Kartenbox füllt.
    s_w = card_w / float(INNER_W_PX)
    s_h = card_h / float(INNER_H_PX)
    s = s_w if abs((s_w * INNER_H_PX) - card_h) <= 0.5 else s_h

    # NEU: Umschließende Kanten aus belegten Zellen ableiten
    min_row, max_row, min_col_row, max_col_row = _compute_enclosing_edges(
        img_paths[:per_page], cols, rows, is_back=is_back
    )

    # NEU: Occupancy-Matrix zur Prüfung „ist die Zelle darunter belegt?“
    occ = [[False] * cols for _ in range(rows)]
    for idx, p in enumerate(img_paths[:per_page]):
        if p and p.exists():
            r = idx // cols
            ccol = idx % cols
            if is_back:
                ccol = (cols - 1) - ccol
            occ[r][ccol] = True

    # Zeichenschleife über alle (theoretischen) Zellen der Seite
    for idx in range(per_page):
        img_path = img_paths[idx] if idx < len(img_paths) else None
        row = idx // cols
        col = idx % cols
        # Rückseite: Spalten spiegeln (Short-edge Duplex Verhalten)
        if is_back:
            col = (cols - 1) - col

        x = x0 + col * card_w
        # y top-down: row==0 visuell OBEN, row==rows-1 UNTEN
        y = y0 + (rows - 1 - row) * card_h

        if img_path is None or not img_path.exists():
            continue

        # --- Außen-Bleed nur an den 'logischen' Rasteraußenkanten ---
        # Links/Rechts: erste/letzte belegte Spalte pro Zeile
        keep_left  = outer_bleed_keep_px if (min_col_row[row] is not None and col == min_col_row[row]) else 0
        keep_right = outer_bleed_keep_px if (max_col_row[row] is not None and col == max_col_row[row]) else 0

        # Oben: erste belegte Zeile insgesamt
        keep_top   = outer_bleed_keep_px if (row == min_row) else 0

        # Unten: entweder letzte belegte Zeile ODER (wenn es eine Zeile darunter gibt)
        # die Zelle direkt darunter ist NICHT belegt → Bleed unten zeichnen.
        if row == max_row:
            keep_bottom = outer_bleed_keep_px
        else:
            if row + 1 < rows:
                # Achtung: 'occ' ist bereits ggf. gespiegelt für Rückseiten
                keep_bottom = outer_bleed_keep_px if not occ[row + 1][col] else 0
            else:
                keep_bottom = 0

        use_outer = outer_bleed_keep_px > 0 and (keep_left or keep_right or keep_top or keep_bottom)

        if use_outer:
            # Nur wenn die Quelle echtes Bleed hat (>= BLEED_W/H_PX); sonst Fallback auf Innenmaß
            sz = get_image_px_size(img_path)
            has_bleed = bool(sz and sz[0] >= BLEED_W_PX and sz[1] >= BLEED_H_PX)
            if has_bleed:
                # Quelle so vorbereiten, dass die Innenfläche erhalten bleibt und außen nur an
                # angegebenen Kanten ein dünner Bleed stehen bleibt.
                processed = preprocess_card_image_outer_bleed(
                    img_path, quality_key,
                    keep_left, keep_right, keep_top, keep_bottom,
                    rotate_degrees=0
                )
                # Gesamtgröße inkl. außenstehender Bleed-Pixel in Punkten
                total_w = s * (INNER_W_PX + keep_left + keep_right)
                total_h = s * (INNER_H_PX + keep_top + keep_bottom)
                # Bild so platzieren, dass die Innenfläche exakt in der Kartenbox liegt
                dx = x - s * keep_left
                dy = y - s * keep_bottom
                # preserveAspectRatio=False, da wir die exakten Maße vorgeben
                c.drawImage(
                    ImageReader(str(processed)),
                    dx, dy,
                    width=total_w, height=total_h,
                    preserveAspectRatio=False, mask="auto"
                )
            else:
                # Fallback: Innenmaß
                processed = preprocess_card_image_for_pdf(img_path, quality_key, card_box_inches)
                draw_w, draw_h = fit_image_into_box(processed, card_w, card_h)
                dx = x + (card_w - draw_w) / 2.0
                dy = y + (card_h - draw_h) / 2.0
                c.drawImage(
                    ImageReader(str(processed)),
                    dx, dy,
                    width=draw_w, height=draw_h,
                    preserveAspectRatio=True, mask="auto"
                )
        else:
            # Kein Außen-Bleed angefragt oder Karte liegt nicht außen → klassisch Innenmaß
            processed = preprocess_card_image_for_pdf(img_path, quality_key, card_box_inches)
            draw_w, draw_h = fit_image_into_box(processed, card_w, card_h)
            dx = x + (card_w - draw_w) / 2.0
            dy = y + (card_h - draw_h) / 2.0
            c.drawImage(
                ImageReader(str(processed)),
                dx, dy,
                width=draw_w, height=draw_h,
                preserveAspectRatio=True, mask="auto"
            )

    # Marken IMMER zeichnen – unabhängig davon, ob außen Bleed stand:
    draw_inner_crosses_grid(c, x0, y0, card_w, card_h, cols, rows)
    draw_outer_marks_grid(c, x0, y0, card_w, card_h, cols, rows)
    draw_corner_marks_grid(c, x0, y0, card_w, card_h, cols, rows)

# Abwärtskompatibel: 3x3 ruft generisch auf
def place_images_3x3(c: canvas.Canvas,
                     img_paths: List[Optional[Path]],
                     x0: float, y0: float,
                     card_w: float, card_h: float,
                     is_back: bool,
                     quality_key: str,
                     card_box_inches: Tuple[float, float]):
    # 9 images, 3 rows x 3 cols
    for idx, img_path in enumerate(img_paths[:9]):
        row = idx // 3
        col = idx % 3

        # Back side: mirror columns (as before)
        if is_back:
            col = 2 - col

        x = x0 + col * card_w
        y = y0 + (2 - row) * card_h

        if img_path is None or not img_path.exists():
            continue

        processed = preprocess_card_image_for_pdf(img_path, quality_key, card_box_inches)
        draw_w, draw_h = fit_image_into_box(processed, card_w, card_h)
        dx = x + (card_w - draw_w) / 2.0
        dy = y + (card_h - draw_h) / 2.0
        c.drawImage(ImageReader(str(processed)), dx, dy, width=draw_w, height=draw_h,
                    preserveAspectRatio=True, mask="auto")

    draw_inner_crosses_3x3(c, x0, y0, card_w, card_h)
    draw_outer_marks_3x3(c, x0, y0, card_w, card_h)
    draw_corner_marks_3x3(c, x0, y0, card_w, card_h)


# =========================================================
# Layout 2x3: landscape, outer cut marks ONLY for poker cutlines
# =========================================================
def get_bleed_box_size_pt() -> Tuple[float, float]:
    """
    The actual poker area is still 2.5"x3.5" (180x252 pt),
    but the full image is 825x1125 px, inner is 750x1050 px.
    So the full box is scaled by (825/750, 1125/1050).
    """
    box_w = POKER_W_PT * (BLEED_W_PX / INNER_W_PX)   # 180 * 1.1 = 198
    box_h = POKER_H_PT * (BLEED_H_PX / INNER_H_PX)   # 252 * 1.071428... = 270
    return box_w, box_h

def get_bleed_box_inches() -> Tuple[float, float]:
    # Inner card size in inches is derived from the selected format.
    w_in = POKER_W_PT / 72.0
    h_in = POKER_H_PT / 72.0
    return w_in * (BLEED_W_PX / INNER_W_PX), h_in * (BLEED_H_PX / INNER_H_PX)

def draw_cutmarks_bleed_outer_only(c: canvas.Canvas,
                                   x0: float, y0: float,
                                   cols: int, rows: int,
                                   box_w: float, box_h: float):
    """
    Draw ONLY outside marks around the raster, but at all poker cutline positions.
    Works for any cols x rows bleed grid (e.g., BLEED).
    """
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_BLEED)
    c.setStrokeColor(CUTMARK_COLOR)
    grid_w = cols * box_w
    grid_h = rows * box_h
    x_left = x0
    x_right = x0 + grid_w
    y_bottom = y0
    y_top = y0 + grid_h

    # Fractions of poker cutlines within ONE bleed box
    fx_left  = BLEED_LEFT_TOP_PX / BLEED_W_PX
    fx_right = (BLEED_LEFT_TOP_PX + INNER_W_PX) / BLEED_W_PX
    fy_bottom = BLEED_RIGHT_BOTTOM_PX / BLEED_H_PX
    fy_top    = (BLEED_RIGHT_BOTTOM_PX + INNER_H_PX) / BLEED_H_PX

    x_cuts = []
    for j in range(cols):
        box_left = x0 + j * box_w
        x_cuts.append(box_left + fx_left  * box_w)
        x_cuts.append(box_left + fx_right * box_w)

    y_cuts = []
    for i in range(rows):
        box_bottom = y0 + i * box_h
        y_cuts.append(box_bottom + fy_bottom * box_h)
        y_cuts.append(box_bottom + fy_top    * box_h)

    L = CUTMARK_LEN_PT_BLEED
    for x in x_cuts:
        c.line(x, y_bottom - L, x, y_bottom)
        c.line(x, y_top,       x, y_top + L)
    for y in y_cuts:
        c.line(x_left - L, y, x_left, y)
        c.line(x_right,   y, x_right + L, y)
    c.restoreState()

def draw_cutmarks_2x3_outer_only(c: canvas.Canvas,
                                 x0: float, y0: float,
                                 cols: int, rows: int,
                                 box_w: float, box_h: float):
    # Weiterleitung auf generalisierte Funktion
    draw_cutmarks_bleed_outer_only(c, x0, y0, cols, rows, box_w, box_h)

def place_images_2x3(c: canvas.Canvas,
                     img_paths: List[Optional[Path]],
                     x0: float, y0: float,
                     box_w: float, box_h: float,
                     is_back: bool,
                     quality_key: str,
                     card_box_inches: Tuple[float, float]):
    place_images_bleed_grid(c, img_paths, x0, y0, box_w, box_h, cols=3, rows=2, is_back=is_back, quality_key=quality_key, card_box_inches=card_box_inches)

def place_images_bleed_grid(c: canvas.Canvas,
                            img_paths: List[Optional[Path]],
                            x0: float, y0: float,
                            box_w: float, box_h: float,
                            cols: int, rows: int,
                            is_back: bool,
                            quality_key: str,
                            card_box_inches: Tuple[float, float]):
    per_page = cols * rows
    for idx, img_path in enumerate(img_paths[:per_page]):
        row = idx // cols
        col = idx % cols
        # Rückseite: Spalten spiegeln (Short-edge Duplex Verhalten wie 2x3)
        if is_back:
            col = (cols - 1) - col
        x = x0 + col * box_w
        y = y0 + (rows - 1 - row) * box_h
        if img_path is None or not img_path.exists():
            continue
        processed = preprocess_card_image_for_pdf(img_path, quality_key, card_box_inches, crop_bleed=False)
        draw_w, draw_h = fit_image_into_box(processed, box_w, box_h)
        dx = x + (box_w - draw_w) / 2.0
        dy = y + (box_h - draw_h) / 2.0
        c.drawImage(ImageReader(str(processed)), dx, dy, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")

    draw_cutmarks_bleed_outer_only(c, x0, y0, cols=cols, rows=rows, box_w=box_w, box_h=box_h)

def place_images_gutterfold_grid(
    c: canvas.Canvas,
    pairs_group: List[Tuple[str, Optional[Path], Optional[Path]]],
    x0: float, y0: float,
    card_w: float, card_h: float,
    cols: int,
    fold_gutter: float,
    quality_key: str,
    card_box_inches: Tuple[float, float],
    outer_bleed_keep_px: int = 0
):
    per_page = cols
    padded = pairs_group + [("", None, None)] * (per_page - len(pairs_group))

    # ============================================================
    # NEU: Ermitteln, welche Spalten tatsächlich belegt sind
    # ============================================================
    used_cols = []
    for col in range(cols):
        _base, front, back = padded[col]
        used_cols.append(
            bool(front and front.exists()) or bool(back and back.exists())
        )

    first_used_col = next(
        (j for j, used in enumerate(used_cols) if used), 0
    )
    last_used_col = next(
        (j for j in range(cols - 1, -1, -1) if used_cols[j]), cols - 1
    )

    # ------------------------------------------------------------
    # Grid-Geometrie
    # ------------------------------------------------------------
    grid_w = cols * card_w
    grid_h = 2 * card_h + fold_gutter

    y_bottom = y0
    y_top = y0 + card_h + fold_gutter
    fold_y = y0 + card_h + fold_gutter / 2.0

    # ------------------------------------------------------------
    # Bildplatzierung
    # ------------------------------------------------------------
    for col in range(cols):
        base, front, back = padded[col]
        x = x0 + col * card_w

        # ---------- FRONT ----------
        if front and front.exists():
            if outer_bleed_keep_px > 0:
                keep_left  = outer_bleed_keep_px if col == first_used_col else 0
                keep_right = outer_bleed_keep_px if col == last_used_col  else 0
                keep_top   = outer_bleed_keep_px
                keep_bottom= outer_bleed_keep_px

                processed_f = preprocess_card_image_outer_bleed(
                    front, quality_key,
                    keep_left, keep_right, keep_top, keep_bottom,
                    rotate_degrees=0
                )

                draw_card_outer_bleed(
                    c, processed_f,
                    x, y_top,
                    card_w, card_h,
                    keep_left, keep_right, keep_top, keep_bottom
                )
            else:
                processed_f = preprocess_card_image_for_pdf(
                    front, quality_key, card_box_inches
                )
                draw_image_transformed(
                    c, processed_f,
                    x, y_top,
                    card_w, card_h,
                    rotate_deg=0,
                    mirror_x=False
                )

        # ---------- BACK ----------
        if back and back.exists():
            if outer_bleed_keep_px > 0:
                keep_left  = outer_bleed_keep_px if col == first_used_col else 0
                keep_right = outer_bleed_keep_px if col == last_used_col  else 0
                keep_top   = outer_bleed_keep_px
                keep_bottom= outer_bleed_keep_px

                processed_b = preprocess_card_image_outer_bleed(
                    back, quality_key,
                    keep_left, keep_right, keep_top, keep_bottom,
                    rotate_degrees=180
                )

                draw_card_outer_bleed(
                    c, processed_b,
                    x, y_bottom,
                    card_w, card_h,
                    keep_left, keep_right, keep_top, keep_bottom
                )
            else:
                processed_b = preprocess_card_image_for_pdf(
                    back, quality_key, card_box_inches
                )
                draw_image_transformed(
                    c, processed_b,
                    x, y_bottom,
                    card_w, card_h,
                    rotate_deg=180,
                    mirror_x=False
                )

    # ------------------------------------------------------------
    # Falzlinie
    # ------------------------------------------------------------
    if GF_DRAW_FOLD_LINE:
        draw_gutterfold_line_horizontal(c, x0, fold_y, grid_w)

    # ------------------------------------------------------------
    # Außenmarken
    # ------------------------------------------------------------
    x_marks = [x0 + j * card_w for j in range(cols + 1)]
    y_edges = sorted({
        y0,
        y0 + card_h,
        y0 + card_h + fold_gutter,
        y0 + grid_h
    })

    draw_cutmarks_gutterfold(
        c,
        x0=x0,
        y0=y0,
        grid_w=grid_w,
        grid_h=grid_h,
        y_edges=y_edges,
        x_marks=x_marks
    )

    # ------------------------------------------------------------
    # Brückenmarken im Gutter
    # ------------------------------------------------------------
    y_gutter_bottom = y0 + card_h
    y_gutter_top    = y0 + card_h + fold_gutter
    bridge_x = [x0 + j * card_w for j in range(cols + 1)]

    draw_gutter_bridge_marks(
        c, bridge_x, y_gutter_bottom, y_gutter_top
    )

# =========================================================
# PDF generation
# =========================================================
def generate_pdf(layout_key: str,
                 out_path: Path,
                 pagesize_tuple: Tuple[float, float],
                 pairs: List[Tuple[str, Optional[Path], Optional[Path]]],
                 logo_path: Optional[Path],
                 copyright_name: Optional[str],
                 version_str: str,
                 quality_key: str,
                 include_back_pages: bool = True,
                 outer_bleed_keep_px: int = 0,
                 rulebook_images: Optional[List[Path]] = None):
    """
    Dynamische Layout-Erzeugung:
      - 'standard' (Innenbilder, ohne Bleed, mit inneren Kreuzen + Außenmarken)
      - 'bleed'    (Bleed-Box, NUR Außenmarken)
      - 'gutterfold' (2 Reihen + Falzgürtel, Brückenmarken)
    Legacy-Keys ('3x3','3x4','2x3','2x5') werden weiter akzeptiert.
    """
    lk = layout_key.strip().lower()
    page_w, page_h = pagesize_tuple

    # --- STANDARD (Innenbilder) ---
    if lk in ("standard", "3x3", "3x4"):
        
        def _compute_header_h_for_logo(logo_path, page_w, page_h, margins, grid_top_y):
            if not logo_path:
                return 0.0
            lw, lh = fit_logo_with_constraints(logo_path, LOGO_MAX_W, LOGO_MAX_H)
            max_header_h = max(0.0, page_h - margins["top"] - grid_top_y - LOGO_GAP_TO_GRID)
            return min(lh, max_header_h)

        card_w, card_h = POKER_W_PT, POKER_H_PT

        # Logo IMMER zeichnen, aber NICHT als harte Reserve in der Platzberechnung
        _apply_logo = bool(logo_path)
        _logo_for_calc = None
        cols, rows, x0, y0, grid_w, grid_h, grid_top_y = compute_max_grid_counts(
            page_w, page_h, card_w, card_h,
            MARGINS_PT, _logo_for_calc, BOTTOM_RESERVED_PT, extra_vertical_pt=0.0
        )
        
        per_page = cols * rows
        c = create_pdf_canvas(out_path, pagesize_tuple, author=(copyright_name or ''))
        draw_rulebook_pages(c, pagesize_tuple, rulebook_images or [], mode="portrait_pref", force_mode=RULEBOOK_ROTATE_MODE)
        # Optionaler Spezialfall: Letter + Standard (früher 3x3 tiefer)
        bottom_y_override = BOTTOM_Y_LETTER_3X3 if pagesize_tuple == letter else None
        
        sheet_no = 0
        for group in chunk(pairs, per_page):
            sheet_no += 1
            fronts = [a for (_n, a, _b) in group] + [None] * (per_page - len(group))
            backs  = [b for (_n, _a, b) in group] + [None] * (per_page - len(group))
            place_images_grid_inner(
                c, fronts, x0, y0, card_w, card_h,
                cols=cols, rows=rows, is_back=False,
                quality_key=quality_key,
                card_box_inches=(POKER_W_PT/72.0, POKER_H_PT/72.0),
                outer_bleed_keep_px=outer_bleed_keep_px
            )
            
            if _apply_logo:
                header_h = _compute_header_h_for_logo(logo_path, page_w, page_h, MARGINS_PT, grid_top_y)
                if header_h > 1.0:
                    draw_logo_in_header_band(c, logo_path, page_w, page_h, MARGINS_PT, header_h)
            draw_bottom_line(c, page_w, copyright_name, version_str, f"{sheet_no}a",
                             y_override=bottom_y_override)
            c.showPage()
            if include_back_pages and any(p for p in backs if p and p.exists()):
                place_images_grid_inner(
                    c, backs, x0 + BACK_X_OFFSET_PT, y0 + BACK_Y_OFFSET_PT, card_w, card_h,
                    cols=cols, rows=rows, is_back=True,
                    quality_key=quality_key,
                    card_box_inches=(POKER_W_PT/72.0, POKER_H_PT/72.0),
                    outer_bleed_keep_px=outer_bleed_keep_px
                )   
                if _apply_logo:
                    header_h = _compute_header_h_for_logo(logo_path, page_w, page_h, MARGINS_PT, grid_top_y)
                    if header_h > 1.0:
                        draw_logo_in_header_band(c, logo_path, page_w, page_h, MARGINS_PT, header_h)

                draw_bottom_line(c, page_w, copyright_name, version_str, f"{sheet_no}b",
                                 y_override=bottom_y_override)
                c.showPage()
        c.save()
        return

    # --- BLEED (nur Außenmarken) ---
    if lk in ("bleed", "2x3", "2x5"):
        
        def _compute_header_h_for_logo(logo_path, page_w, page_h, margins, grid_top_y):
            if not logo_path:
                return 0.0
            lw, lh = fit_logo_with_constraints(logo_path, LOGO_MAX_W, LOGO_MAX_H)
            max_header_h = max(0.0, page_h - margins["top"] - grid_top_y - LOGO_GAP_TO_GRID)
            return min(lh, max_header_h)
        
        box_w, box_h = get_bleed_box_size_pt()
        # Logo IMMER zeichnen, aber NICHT als harte Reserve
        _apply_logo = bool(logo_path)
        _logo_for_calc = None
        cols, rows, x0, y0, grid_w, grid_h, grid_top_y = compute_max_grid_counts(
            page_w, page_h, box_w, box_h,
            MARGINS_PT, _logo_for_calc, BOTTOM_RESERVED_PT, extra_vertical_pt=0.0
        )
        per_page = cols * rows
        c = create_pdf_canvas(out_path, pagesize_tuple, author=(copyright_name or ''))
        draw_rulebook_pages(c, pagesize_tuple, rulebook_images or [], mode="landscape_pref", force_mode=RULEBOOK_ROTATE_MODE)
        sheet_no = 0
        for group in chunk(pairs, per_page):
            sheet_no += 1
            fronts = [a for (_n, a, _b) in group] + [None] * (per_page - len(group))
            backs  = [b for (_n, _a, b) in group] + [None] * (per_page - len(group))
            place_images_bleed_grid(
                c, fronts, x0, y0, box_w, box_h,
                cols=cols, rows=rows, is_back=False,
                quality_key=quality_key,
                card_box_inches=get_bleed_box_inches()
            )


            if _apply_logo:
                header_h = _compute_header_h_for_logo(logo_path, page_w, page_h, MARGINS_PT, grid_top_y)
                if header_h > 1.0:
                    draw_logo_in_header_band(c, logo_path, page_w, page_h, MARGINS_PT, header_h)
            draw_bottom_line(c, page_w, copyright_name, version_str, f"{sheet_no}a")
            c.showPage()
            if include_back_pages and any(p for p in backs if p and p.exists()):
                place_images_bleed_grid(
                    c, backs, x0 + BACK_X_OFFSET_PT, y0 + BACK_Y_OFFSET_PT, box_w, box_h,
                    cols=cols, rows=rows, is_back=True,
                    quality_key=quality_key,
                   card_box_inches=get_bleed_box_inches()
                )
                if _apply_logo:
                    header_h = _compute_header_h_for_logo(logo_path, page_w, page_h, MARGINS_PT, grid_top_y)
                    if header_h > 1.0:
                        draw_logo_in_header_band(c, logo_path, page_w, page_h, MARGINS_PT, header_h)
                draw_bottom_line(c, page_w, copyright_name, version_str, f"{sheet_no}b")
                c.showPage()
        c.save()
        return

    # --- GUTTERFOLD ---
    if lk in ("gutterfold",):
        
        def _compute_header_h_for_logo(logo_path, page_w, page_h, margins, grid_top_y):
                if not logo_path:
                    return 0.0
                lw, lh = fit_logo_with_constraints(logo_path, LOGO_MAX_W, LOGO_MAX_H)
                max_header_h = max(0.0, page_h - margins["top"] - grid_top_y - LOGO_GAP_TO_GRID)
                return min(lh, max_header_h)
        
        card_w, card_h = POKER_W_PT, POKER_H_PT
        gf_extra = GF_FOLD_GUTTER_PT
        # 2 Reihen fix; Spalten dynamisch:
        # Logo nicht als harte Reserve; Kopfband nur über RESERVE_TOP_PT
        top_res = RESERVE_TOP_PT
        avail_w = page_w - MARGINS_PT["left"] - MARGINS_PT["right"]
        avail_h = page_h - MARGINS_PT["top"] - MARGINS_PT["bottom"] - top_res - BOTTOM_RESERVED_PT
        if avail_h < (2 * card_h + gf_extra):
            raise ValueError(
                t(
                    "error_gutterfold_space",
                    avail=avail_h,
                    need=needed_h,
                    margin=PRINT_SAFE_MARGIN_CM,
                    top=top_res,
                    bottom=BOTTOM_RESERVED_PT
                )
            )
        cols = max(1, int(avail_w // card_w))
        grid_w = cols * card_w
        grid_h = 2 * card_h + gf_extra
        x0, y0 = compute_grid_origin_centered_with_margins(
            page_w, page_h, grid_w, grid_h,
            MARGINS_PT, top_res, BOTTOM_RESERVED_PT
        )
        grid_top_y = y0 + grid_h
        per_page = cols  # je Spalte ein Paar
        c = create_pdf_canvas(out_path, pagesize_tuple, author=(copyright_name or ''))
        draw_rulebook_pages(c, pagesize_tuple, rulebook_images or [], mode="landscape_pref", force_mode=RULEBOOK_ROTATE_MODE)
        _apply_logo = bool(logo_path)
        sheet_no = 0
        for group in chunk(pairs, per_page):
            sheet_no += 1
            place_images_gutterfold_grid(
                c, group, x0, y0, card_w, card_h,
                cols=cols, fold_gutter=GF_FOLD_GUTTER_PT,
                quality_key=quality_key,
                card_box_inches=(POKER_W_PT/72.0, POKER_H_PT/72.0),
                outer_bleed_keep_px=outer_bleed_keep_px
            )
            
            if _apply_logo:
                header_h = _compute_header_h_for_logo(logo_path, page_w, page_h, MARGINS_PT, grid_top_y)
                if header_h > 1.0:
                    draw_logo_in_header_band(c, logo_path, page_w, page_h, MARGINS_PT, header_h)
            draw_bottom_line(c, page_w, copyright_name, version_str, f"{sheet_no}")
            c.showPage()
        c.save()
        return

    raise ValueError("Unknown layout_key")

def choose_gutterfold_orientation(base_pagesize):
    """Wählt die Orientierung (landscape/portrait), die 2 Reihen + Falzgürtel in Originalgröße erlaubt."""
    def available_h(pagesize_tuple):
        page_w, page_h = pagesize_tuple
        return (page_h 
                - MARGINS_PT["top"] - MARGINS_PT["bottom"] 
                - RESERVE_TOP_PT - BOTTOM_RESERVED_PT)

    needed_h = 2 * POKER_H_PT + GF_FOLD_GUTTER_PT  # Kartenhöhe stammt aus dem gewählten Format
    ls = landscape(base_pagesize)
    if available_h(ls) >= needed_h:
        return ls  # Querformat passt
    if available_h(base_pagesize) >= needed_h:
        return base_pagesize  # Hochformat passt
    # Falls beides nicht reicht, nimm die Variante mit mehr Höhe; generate_pdf wird trotzdem sauber aborten.
    return ls if available_h(ls) >= available_h(base_pagesize) else base_pagesize


# =========================================================
# Main
# =========================================================
def main():
    # -----------------------------
    # 1) Sprache/Start + Header
    # -----------------------------
    args = parse_args()
    # --- Immer beim Start: pdfConfig.txt im App-Ordner anlegen, falls (noch) nicht vorhanden ---
    try:
        cfg_default_path = get_writable_base_dir() / PDF_CONFIG_NAME_DEFAULT
        if not cfg_default_path.exists():
            write_pdf_config_template(cfg_default_path)
    except Exception:
        # Template-Erzeugung darf den Start nicht verhindern
        pass
    if getattr(args, "lang", None):  # CLI-Sprache persistieren
        save_lang_to_ini(args.lang)
    prompt_language_if_needed()  # lädt I18N + INI
    clear_tmp_cache()
    _show_header()                             # rich-Header (oder print)
    print(t("startup_license"))
    print(" ")

    # -----------------------------
    # 2) ZUERST: Kartenordner abfragen
    # -----------------------------
    # (Pfad steht ab hier fest, aber noch KEINE Analysen/Warnungen ausführen.)
    if getattr(args, "folder", None):
        folder = Path(expanduser(args.folder)).expanduser().resolve()
        if not (folder.exists() and folder.is_dir()):
            if rprint:
                rprint(Panel("Ungültiger Ordner. Bitte erneut wählen.", title="Fehler", border_style="red"))
            else:
                print("Ungültiger Ordner. Bitte erneut wählen.")
            folder = prompt_folder()
    else:
        folder = prompt_folder()

    # -----------------------------
    # 2b) NEU: pdfConfig.txt im Kartenordner einlesen (wenn vorhanden)
    # -----------------------------
    cfg_txt = folder / PDF_CONFIG_NAME_DEFAULT
    cfg = read_pdf_config(cfg_txt) if cfg_txt.exists() and cfg_txt.is_file() else {}
    use_cfg = bool(cfg)
    # Konsolenhinweis: pdfConfig.txt wird verwendet (mehrsprachig)
    if use_cfg:
        msg = t("using_pdfconfig", path=str(cfg_txt))
        try:
            if rprint and Panel:
                rprint(Panel.fit(
                    msg,
                    title=t("config_title"),
                    border_style="green"
                ))
            else:
                print(msg)
        except Exception:
            print(msg)

    # -----------------------------
    # 3) DANN: Kartenformat (aus Config oder via Prompt)
    # -----------------------------
    def _prompt_card_format(args=None) -> dict:
        # 1) CLI-Override
        if args and getattr(args, "card_format", None):
            wanted = args.card_format.strip().lower()
            fmt = next((f for f in CARD_FORMATS if f['name'].lower() == wanted), None)
            if fmt:
                return fmt
            print(t('invalid_card_format'))
        # 2) Komfort: questionary
        if questionary is not None:
            q_title = t('choose_card_format')
            
            # Für INI-Custom-Formate (f.get('src') == 'ini') stets auf 1 Nachkommastelle runden,
            # um Darstellungen wie 88.89999999999999 mm zu vermeiden.
            def _choice_mm_pair(f):
                w = float(f['w_mm']); h = float(f['h_mm'])
                if f.get('src') == 'ini':
                    return f"{_mm_str_custom(w)} x {_mm_str_custom(h)}"
                return f"{_mm_str(w)} x {_mm_str(h)}"
    
            choices = [
                f"{f['name']} ({_choice_mm_pair(f)} mm)"
                for f in CARD_FORMATS
            ]    
            picked = _q_select(q_title, choices=choices, default=choices[0])
            try:
                idx = choices.index(picked)
            except Exception:
                idx = 0
            return CARD_FORMATS[idx]
        # 3) Fallback: bestehende Funktion (dein alter Prompt)
        return prompt_card_format()

    if use_cfg:
        # CARD_FORMAT aus cfg (ID -> Dict)
        try:
            wanted_id = int(cfg.get("CARD_FORMAT", "1"))
        except Exception:
            wanted_id = 1
        fmt = next((f for f in CARD_FORMATS if int(f.get('id', -1)) == wanted_id), CARD_FORMATS[0])
    else:
        fmt = _prompt_card_format(args)    
    apply_card_format(fmt)
    print_selected_format_info(fmt)
    _show_format_table(fmt)  # rich-Tabelle + Hinweis-Panel (falls rich vorhanden)

    # -----------------------------
    # 4) JETZT: Layout + Papierformat (aus Config oder via Prompt)
    # -----------------------------
    if use_cfg:
        layout_keys = _map_layout_value(cfg.get("LAYOUT", "All"))
        size_modes  = _map_paper_value(cfg.get("PAPER", "Both"))
    else:
        layout_keys = prompt_layout_dynamic(args)  # ["standard", ...]
        size_modes  = prompt_pagesize_mode(args)   # [(A4,"_A4"), ...]

    # -----------------------------
    # 5) Qualität (aus Config oder via Prompt)
    # -----------------------------
    quality_key = _map_quality_value(cfg.get("QUALITY", "")) if use_cfg else prompt_quality(args)

    # Jetzt erst Paare suchen – der Ordner ist garantiert gesetzt
    pairs = find_card_pairs(folder)
    if not pairs:
        # Kein erneutes Nachfragen: Fehlermeldung zeigen und dann per Enter beenden.
        msg = _build_no_cards_message(folder)
        _show_panel(msg, title=t("no_cards_title"), border_style="red")
        pause_before_exit(t("exit_press_enter"), print_message=True)
        return

    # -----------------------------
    # 6) Rückseiten-Handling, Analyse, Bleed-Checks
    #    (deine bestehende Logik – unverändert übernommen)
    # -----------------------------
    include_back_pages = True
    missing_backs = [base for (base, a, b) in pairs if not (b and b.exists())]
    has_any_back  = len(missing_backs) < len(pairs)

    if not has_any_back:
        shared_back = find_named_image_in_folder(folder, CARDBACK_BASENAME)
        if shared_back:
            print(t('using_cardback', file=shared_back.name))
            pairs = [(base, a, shared_back) for (base, a, _b) in pairs]
            missing_backs = []
            has_any_back = True
        else:
            include_back_pages = False
            if any(k.lower() == "gutterfold" for k in layout_keys):
                layout_keys = [k for k in layout_keys if k.lower() != "gutterfold"]
                print(t('skip_gutterfold_no_backs', name=CARDBACK_BASENAME))

    if missing_backs and any(k.lower() == "gutterfold" for k in layout_keys):
        missing_front_files = []
        for (base, a, b) in pairs:
            if not (b and b.exists()):
                missing_front_files.append(a.name if (a and a.exists()) else str(base))
        missing_sorted = sorted(missing_front_files, key=lambda s: s.lower())
        if len(missing_sorted) > 30:
            shown = ', '.join(missing_sorted[:30]) + f" ... (+{len(missing_sorted)-30})"
        else:
            shown = ', '.join(missing_sorted)
        print(t('skip_gutterfold_missing_backs', missing=shown))
        layout_keys = [k for k in layout_keys if k.lower() != "gutterfold"]

    # Bildanalyse / Eligibility für Bleed-Layouts
    sizes, too_small, too_small_bleed, pairs_for_2x3, skipped_2x3 = analyze_card_images(pairs)

    if too_small:
        names = sorted({p.name for p in too_small})
        msg = t('warn_too_small_upscale', minw=INNER_W_PX, minh=INNER_H_PX, count=len(names), files=', '.join(names))
        if rprint:
            rprint(Panel(msg, title="Warnung: Upscaling", border_style="yellow"))
        else:
            print(msg)

    requested_bleed_any = any(k.lower() in ("bleed", "2x3", "2x5") for k in layout_keys)
    if requested_bleed_any and too_small_bleed:
        names = sorted({p.name for p in too_small_bleed})
        shown = ', '.join(names[:30]) + (f" ... (+{len(names)-30})" if len(names) > 30 else "")
        msg = t("skip_bleed_due_to_small", minw=BLEED_W_PX, minh=BLEED_H_PX, count=len(names), files=shown)
        print(msg)
        layout_keys = [k for k in layout_keys if k.lower() not in ("bleed", "2x3", "2x5")]

    requested_2x3 = any(k.lower() == "2x3" for k in layout_keys)
    requested_2x5 = any(k.lower() == "2x5" for k in layout_keys)
    if (requested_2x3 or requested_2x5) and (not pairs_for_2x3 or len(pairs_for_2x3) != len(pairs)):
        if requested_2x3:
            layout_keys = [k for k in layout_keys if k.lower() != "2x3"]
            print(t('skip_2x3', minw=BLEED_W_PX, minh=BLEED_H_PX))
        if requested_2x5:
            layout_keys = [k for k in layout_keys if k.lower() != "2x5"]
            print(t('skip_2x5', minw=BLEED_W_PX, minh=BLEED_H_PX))

    if not layout_keys:
        pause_before_exit()
        return

    # 7) Rulebook/Logo früh ankündigen, dann weitere Eingaben
    # -------------------------------------------------------
    # Rulebook-Bilder finden & Nutzer informieren
    rulebook_images = find_rulebook_images(folder, RULEBOOK_BASENAME)
    if rulebook_images:
        names_str = ", ".join(p.name for p in rulebook_images[:30]) + (f" ... (+{len(rulebook_images)-30})" if len(rulebook_images) > 30 else "")
        print(t("rulebook_found", files=names_str))
        print(t("rulebook_will_prepend"))
    else:
        print(t("rulebook_not_found", name=RULEBOOK_BASENAME))

    # Logo automatisch suchen & Nutzer informieren
    logo_path = find_named_image_in_folder(folder, LOGO_BASENAME)
    if logo_path:
        print(t("logo_found", file=logo_path.name))
    else:
        print(t("logo_not_found", name=LOGO_BASENAME))

    # 8) Weitere Eingaben (Urheber/Version/Ausgabename) – aus Config oder via Prompt
    # ----------------------------------------------------------
    # (Logo bereits ermittelt; rulebook_images ebenfalls vorhanden)
    if use_cfg:
        # BOTTOM_TEXT
        raw_bottom = cfg.get("BOTTOM_TEXT", "")
        bottom_txt = raw_bottom.replace("(C)", "©").replace("(c)", "©") if raw_bottom else None
        # VERSION
        version_str = cfg.get("VERSION", "").strip()
        # OUTPUT_NAME
        out_base = make_safe_name(cfg.get("OUTPUT_NAME", "cards"))
        # Für Konsistenz mit existierenden Variablennamen:
        copyright_name = (bottom_txt or None)
        try:
            # Layouts anzeigen: "Standard, Bleed, Gutterfold"
            def _cap(s: str) -> str:
                return s[:1].upper() + s[1:] if s else s
            layouts_disp = ", ".join(_cap(k) for k in layout_keys)
            # Paper anzeigen: Both | A4 | Letter
            if len(size_modes) >= 2:
                paper_disp = "Both"
            else:
                paper_disp = "A4" if (size_modes and len(size_modes[0]) > 1 and size_modes[0][1] == "_A4") else "Letter"
            # Quality anzeigen: Lossless|High|Medium|Low
            _qmap = {"lossless": "Lossless", "high": "High", "medium": "Medium", "low": "Low"}
            quality_disp = _qmap.get(quality_key, quality_key)
            # Optionalfelder anzeigen
            text_disp = (bottom_txt.strip() if (bottom_txt and bottom_txt.strip()) else t("none"))
            version_disp = (version_str if version_str else t("none"))
            # Kartenformat (ID + Name)
            fmt_id = str(fmt.get("id", ""))
            fmt_name = str(fmt.get("name", ""))
            summary_lines = [
                t("cfg_format", id=fmt_id, name=fmt_name),
                t("cfg_layouts", layouts=layouts_disp),
                t("cfg_paper", paper=paper_disp),
                t("cfg_quality", quality=quality_disp),
                t("cfg_bottom_text", text=text_disp),
                t("cfg_version", version=version_disp),
                t("cfg_output_name", name=out_base),
            ]
            summary_msg = "\n".join(summary_lines)
            if rprint and Panel:
                rprint(Panel.fit(summary_msg, title=t("config_title"), border_style="cyan"))
            else:
                print(summary_msg)
        except Exception:
            # Bei jeglichem Fehler die Generierung nicht blockieren
            pass
    else:
        copyright_name = getattr(args, "copyright", None)
        if copyright_name is None:
            copyright_name = prompt_copyright_name()
        version_str = getattr(args, "version", None) or prompt_version()
        out_base = getattr(args, "out_base", None) or prompt_output_base("cards")    
    generation_dir = build_generation_dir(out_base)
    
    # -----------------------------
    # 9) Warm-Up (optional, beschleunigt das spätere Zeichnen)
    # -----------------------------
    if "standard" in [k.lower() for k in layout_keys]:
        all_imgs_std = _collect_all_images_for("standard", pairs)
        warmup_preprocessing(all_imgs_std, quality_key, (POKER_W_PT/72.0, POKER_H_PT/72.0), crop_bleed=True)

    if any(k.lower() in ("bleed", "2x3", "2x5") for k in layout_keys):
        all_imgs_bleed = _collect_all_images_for("bleed", pairs)
        warmup_preprocessing(all_imgs_bleed, quality_key, get_bleed_box_inches(), crop_bleed=False)

    # -----------------------------
    # 10) PDF-Erzeugung (wie bisher)
    # -----------------------------
    all_have_bleed = (pairs_for_2x3 and len(pairs_for_2x3) == len(pairs))

    for layout_key in layout_keys:
        # Paare je nach Modus
        if layout_key in ("bleed", "2x3", "2x5"):
            pairs_layout = pairs_for_2x3
        else:
            pairs_layout = pairs

        for base_pagesize, suffix in size_modes:
            # Orientierung
            if layout_key in ("gutterfold",):
                pagesize_tuple = choose_gutterfold_orientation(base_pagesize)
            elif layout_key in ("bleed", "2x3", "2x5"):
                pagesize_tuple = landscape(base_pagesize)
            else:
                pagesize_tuple = base_pagesize

            # Suffix/Outer-bleed pro Layout
            if layout_key in ("standard", "3x3", "3x4"):
                layout_suffix = "_standard"
                outer_keep = OUTER_BLEED_KEEP_PX
            elif layout_key in ("bleed", "2x3", "2x5"):
                layout_suffix = "_bleed"
                outer_keep = 0
            else:
                layout_suffix = "_gutterfold"
                outer_keep = OUTER_BLEED_KEEP_PX if all_have_bleed else 0
               
            # Ausgabepfad: Immer in den Generierungs-Unterordner schreiben
            # Wunsch: erst Papierformat, dann Layout  ? {out_base}{suffix}{layout_suffix}.pdf
            out_path = (generation_dir / f"{out_base}{suffix}{layout_suffix}.pdf").resolve()
            
            # Erzeugung
            generate_pdf(
                layout_key=layout_key,
                out_path=out_path,
                pagesize_tuple=pagesize_tuple,
                pairs=pairs_layout,
                logo_path=logo_path,
                copyright_name=copyright_name,
                version_str=version_str,
                quality_key=quality_key,
                outer_bleed_keep_px=outer_keep,
                rulebook_images=rulebook_images
            )
            print(t("done", path=out_path))

if __name__ == "__main__":
    try:
        main()
        # Only pause once at the very end, and only if we did NOT already pause earlier
        # (e.g. because we exited early due to "No cards found").
        if not _PAUSE_ALREADY_SHOWN:
            pause_before_exit("\n" + t("exit_ok"))
    except Exception:
        print("\n" + t("exit_err_header") + "\n")
        import traceback
        traceback.print_exc()
        # If an earlier pause already happened, don't force a second Enter.
        if not _PAUSE_ALREADY_SHOWN:
            pause_before_exit("\n" + t("exit_press_enter"))
