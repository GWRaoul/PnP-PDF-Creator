# -*- coding: utf-8 -*-
"""
PnP PDF Creator
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
- Pillow / PIL fork (HPND license)

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
import tempfile
import hashlib
import sys
import configparser
from pathlib import Path
from typing import Dict, List, Tuple, Optional

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, letter, landscape
from reportlab.lib.utils import ImageReader

try:
    from PIL import Image
except ImportError:
    Image = None

# =========================================================
# Script version / debug
# =========================================================
SCRIPT_VERSION = 'V1.1-2026-01-28'
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
        print(f"{f['name']} ({w_str} x {h_str} mm) [{f['id']}]")

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
LOGO_MAX_W = 100.0
LOGO_MAX_H = 15.0
LOGO_GAP_TO_GRID = 2.0

# Bottom line
COPY_MAX_CHARS = 30

# 3x3 + Gutterfold cut marks (standard)
# These are overridable via INI (section [cutmarks]).
CUTMARK_LEN_PT_STD = 10.0
CUTMARK_LINE_PT_STD = 1.0

# 2x3 marks (outer only, cut to poker area inside bleed image)
CUTMARK_LEN_PT_2X3 = 20.0
CUTMARK_LINE_PT_2X3 = 1.0
# 2x3 card image geometry (pixels of the source image)
BLEED_W_PX = 825
BLEED_H_PX = 1125
BLEED_LEFT_TOP_PX = 37
BLEED_RIGHT_BOTTOM_PX = 38
INNER_W_PX = BLEED_W_PX - BLEED_LEFT_TOP_PX - BLEED_RIGHT_BOTTOM_PX  # 750
INNER_H_PX = BLEED_H_PX - BLEED_LEFT_TOP_PX - BLEED_RIGHT_BOTTOM_PX  # 1050

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

def get_app_dir() -> Path:
    """
    Directory where the EXE resides (PyInstaller) or script directory (normal python).
    For PyInstaller --onefile, sys.executable points to the actual .exe path.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def get_ini_path() -> Path:
    return get_app_dir() / "PnP_PDF_Creator.ini"

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


def ensure_cutmark_defaults(cp: configparser.ConfigParser) -> bool:
    # Ensure [cutmarks] section exists with defaults. Returns True if cp was modified.
    changed = False
    if not cp.has_section('cutmarks'):
        cp.add_section('cutmarks')
        changed = True

    defaults = {
        'length_pt_standard': str(CUTMARK_LEN_PT_STD),
        'width_pt_standard': str(CUTMARK_LINE_PT_STD),
        'length_pt_2x3': str(CUTMARK_LEN_PT_2X3),
        'width_pt_2x3': str(CUTMARK_LINE_PT_2X3),
    }
    for k, v in defaults.items():
        if not cp.has_option('cutmarks', k):
            cp.set('cutmarks', k, v)
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
    # Ensure [assets] section exists with defaults (e.g. shared cardback image name).
    changed = False
    if not cp.has_section('assets'):
        cp.add_section('assets')
        changed = True
    if not cp.has_option('assets', 'cardback_name'):
        cp.set('assets', 'cardback_name', DEFAULT_CARDBACK_BASENAME)
        changed = True
    return changed


def load_assets_from_config(cp: configparser.ConfigParser) -> None:
    # Load asset settings from INI into global variables.
    global CARDBACK_BASENAME
    name = cp.get('assets', 'cardback_name', fallback=DEFAULT_CARDBACK_BASENAME).strip()
    CARDBACK_BASENAME = name if name else DEFAULT_CARDBACK_BASENAME
def load_cutmarks_from_config(cp: configparser.ConfigParser) -> None:
    # Load cutmark settings from INI into the global variables.
    global CUTMARK_LEN_PT_STD, CUTMARK_LINE_PT_STD, CUTMARK_LEN_PT_2X3, CUTMARK_LINE_PT_2X3

    CUTMARK_LEN_PT_STD = _get_positive_float(cp, 'cutmarks', 'length_pt_standard', CUTMARK_LEN_PT_STD)
    CUTMARK_LINE_PT_STD = _get_positive_float(cp, 'cutmarks', 'width_pt_standard', CUTMARK_LINE_PT_STD)
    CUTMARK_LEN_PT_2X3 = _get_positive_float(cp, 'cutmarks', 'length_pt_2x3', CUTMARK_LEN_PT_2X3)
    CUTMARK_LINE_PT_2X3 = _get_positive_float(cp, 'cutmarks', 'width_pt_2x3', CUTMARK_LINE_PT_2X3)

def save_lang_to_ini(lang: str) -> None:
    cp = load_config()
    if not cp.has_section('ui'):
        cp.add_section('ui')
    cp.set('ui', 'lang', lang)
    # Ensure cutmark defaults exist so users can edit them
    ensure_cutmark_defaults(cp)
    # Ensure assets defaults exist so users can edit them
    ensure_assets_defaults(cp)
    write_config(cp)

def prompt_language_if_needed():
    global LANG
    cp = load_config()
    changed = ensure_cutmark_defaults(cp)
    changed = ensure_assets_defaults(cp) or changed
    changed = ensure_custom_format_defaults(cp) or changed

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

    # Persist INI if defaults were added or language was set
    if changed:
        write_config(cp)

I18N = {
    "de": {
        "choose_layout": "Layout wählen ({opts}) [All]: ",
        "invalid_layout": "Bitte eines der angebotenen Layouts eingeben.",
        "format_info_note": "Hinweis: 'Standard' verwendet Innenbilder ohne Beschnitt. 'Bleed' benötigt Bilder mit Beschnitt. 'Gutterfold' erstellt ein Falzlayout mit vorderer/hinterer Seite.",
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
        "ask_logo": "Optional: Pfad zur Logo-Datei (Enter = auto: logo.png/jpg im Kartenordner): ",
        "logo_invalid": "Logo-Pfad ungueltig oder Datei nicht gefunden. Es wird ohne Logo fortgefahren.",
        "ask_quality": "Qualitaet waehlen (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Ungueltige Qualitaet. Bitte Lossless, High, Medium oder Low eingeben.",
        "ask_copyright": "Copyright einbauen? Name eingeben (Enter = nein): ",
        "ask_version": "Versionsnummer eingeben (Enter = leer): ",
        "ask_out_base": "Ausgabedatei Basisname (ohne .pdf) [{default}]: ",
        "no_cards": "Keine Karten gefunden. Erwartet: Dateinamen enden auf 'a' oder 'b' (z.B. card01a.png / card01b.png) ODER enden auf '[face,<n>]' bzw. '[back,<n>]' (z.B. card01[face,001].png / card01[back,001].png).",
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
    },
    "en": {
        "choose_layout": "Choose layout ({opts}) [All]: ",
        "invalid_layout": "Please enter one of the offered layouts.",
        "format_info_note": "Note: 'Standard' uses inner images (no bleed). 'Bleed' requires bleed images. 'Gutterfold' produces a fold layout with matching front/back alignment.",
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
        "ask_logo": "Optional: path to logo file (Enter = auto: logo.png/jpg in card folder): ",
        "logo_invalid": "Invalid logo path or file not found. Continuing without logo.",
        "ask_quality": "Choose quality (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Invalid quality. Please enter Lossless, High, Medium or Low.",
        "ask_copyright": "Add copyright? Enter name (Enter = no): ",
        "ask_version": "Enter version string (Enter = empty): ",
        "ask_out_base": "Output base filename (without .pdf) [{default}]: ",
        "no_cards": "No cards found. Expected filenames ending with 'a' or 'b' (e.g. card01a.png / card01b.png) OR ending with '[face,<n>]' / '[back,<n>]' (e.g. card01[face,001].png / card01[back,001].png).",
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
    },
    "fr": {
        "choose_layout": "Choisissez un layout ({opts}) [All] : ",
        "invalid_layout": "Veuillez saisir l’un des layouts proposés.",
        "format_info_note": "Remarque : 'Standard' utilise des images internes sans fond perdu. 'Bleed' nécessite des images avec fond perdu. 'Gutterfold' crée une mise en page pliée avec alignement recto/verso.",
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
        "ask_logo": "Optionnel : chemin du logo (Entrer = auto : logo.png/jpg dans le dossier) : ",
        "logo_invalid": "Chemin du logo invalide ou fichier introuvable. Suite sans logo.",
        "ask_quality": "Choisir la qualite (Lossless/High/Medium/Low) [High] : ",
        "invalid_quality": "Qualite invalide. Entrez Lossless, High, Medium ou Low.",
        "ask_copyright": "Ajouter un copyright ? Entrez un nom (Entrer = non) : ",
        "ask_version": "Entrer la version (Entrer = vide) : ",
        "ask_out_base": "Nom de fichier de sortie (sans .pdf) [{default}] : ",
        "no_cards": "Aucune carte trouvée. Attendu : noms finissant par 'a' ou 'b' (ex. card01a.png / card01b.png) OU finissant par '[face,<n>]' / '[back,<n>]' (ex. card01[face,001].png / card01[back,001].png).",
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
    },
    "es": {
        "choose_layout": "Elija un layout ({opts}) [All]: ",
        "invalid_layout": "Introduzca uno de los layouts ofrecidos.",
        "format_info_note": "Nota: 'Standard' utiliza imágenes internas sin sangrado. 'Bleed' requiere imágenes con sangrado. 'Gutterfold' crea un diseño plegado con alineación anverso/reverso.",
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
        "ask_logo": "Opcional: ruta del logo (Enter = auto: logo.png/jpg en la carpeta): ",
        "logo_invalid": "Ruta de logo inválida o archivo no encontrado. Continuando sin logo.",
        "ask_quality": "Elegir calidad (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Calidad inválida. Introduce Lossless, High, Medium o Low.",
        "ask_copyright": "¿Agregar copyright? Introduce nombre (Enter = no): ",
        "ask_version": "Introduce versión (Enter = vacío): ",
        "ask_out_base": "Nombre base de salida (sin .pdf) [{default}]: ",
        "no_cards": "No se encontraron cartas. Se esperan nombres que terminen en 'a' o 'b' (p.ej. card01a.png / card01b.png) O que terminen en '[face,<n>]' / '[back,<n>]' (p.ej. card01[face,001].png / card01[back,001].png).",
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
    },
    "it": {
        "choose_layout": "Scegli un layout ({opts}) [All]: ",
        "invalid_layout": "Inserire uno dei layout proposti.",
        "format_info_note": "Nota: 'Standard' utilizza immagini interne senza abbondanza. 'Bleed' richiede immagini con abbondanza. 'Gutterfold' crea un layout piegato con allineamento fronte/retro.",
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
        "ask_logo": "Opzionale: percorso logo (Invio = auto: logo.png/jpg nella cartella): ",
        "logo_invalid": "Percorso logo non valido o file non trovato. Continuo senza logo.",
        "ask_quality": "Scegli qualita (Lossless/High/Medium/Low) [High]: ",
        "invalid_quality": "Qualita non valida. Inserisci Lossless, High, Medium o Low.",
        "ask_copyright": "Inserire copyright? Nome (Invio = no): ",
        "ask_version": "Inserisci versione (Invio = vuoto): ",
        "ask_out_base": "Nome base output (senza .pdf) [{default}]: ",
        "no_cards": "Nessuna carta trovata. Atteso: nomi che terminano con 'a' o 'b' (es. card01a.png / card01b.png) O che terminano con '[face,<n>]' / '[back,<n>]' (es. card01[face,001].png / card01[back,001].png).",
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
    },
}

def t(key: str, **kwargs) -> str:
    lang = I18N.get(LANG, I18N["de"])
    msg = lang.get(key, I18N["de"].get(key, key))
    return msg.format(**kwargs)

# =========================================================
# Console pause helper (useful for PyInstaller EXE)
# =========================================================
def pause_before_exit(message: str = "") -> None:
    """Wait for Enter so the console window stays open (mainly for EXE runs)."""
    try:
        if message:
            print(message)
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
        # Cannot inspect sizes without PIL -> safest behavior: disable 2x3 and skip size-based cropping checks.
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

    # check minimum size and build 2x3 eligible list
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
def prompt_layout_dynamic() -> List[str]:
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


def prompt_pagesize_mode():
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
        p = input(t("ask_folder")).strip().strip('"')
        folder = Path(p)
        if folder.exists() and folder.is_dir():
            return folder
        print(t("invalid_folder"))

def prompt_logo_path(folder: Path) -> Optional[Path]:
    p = input(t("ask_logo")).strip().strip('"')

    if p:
        lp = Path(p)
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

def prompt_quality() -> str:
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
    name = input(t("ask_copyright")).strip()
    if not name:
        return None
    return name[:COPY_MAX_CHARS]

def prompt_version() -> str:
    return input(t("ask_version")).strip()

def prompt_output_base(default_base: str) -> str:
    base = input(t("ask_out_base", default=default_base)).strip()
    return base if base else default_base


# =========================================================
# Card pairing
# =========================================================
def find_card_pairs(folder: Path) -> List[Tuple[str, Optional[Path], Optional[Path]]]:
    """Find and pair card front/back images.

    Supported naming schemes (case-insensitive):

    1) Legacy suffix scheme:
       - Front ends with 'a'  (e.g. card01a.png)
       - Back  ends with 'b'  (e.g. card01b.png)

    2) Bracket scheme:
       - Front contains '[face,<n>]' at the END of the filename stem
       - Back  contains '[back,<n>]' at the END of the filename stem
       where <n> can be 1-3 digits (e.g. [face,1], [face,001], [back,123]).

    Pairing key is the base name + the bracket number (if present).
    """
    files = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXT]

    # scheme 1: ...a / ...b
    ab_pattern = re.compile(r"^(.*)([ab])$", re.IGNORECASE)
    # scheme 2: ...[face,001] / ...[back,001]
    bracket_pattern = re.compile(r"^(.*)\[(face|back),(\d{1,3})\]$", re.IGNORECASE)

    # key -> {'a': Path, 'b': Path, 'base': str, 'num': Optional[str]}
    pairs: Dict[str, Dict[str, object]] = {}

    for f in files:
        stem = f.stem
        m2 = bracket_pattern.match(stem)
        if m2:
            base = m2.group(1)
            kind = m2.group(2).lower()  # face/back
            num = m2.group(3)
            num_norm = num.zfill(3)  # normalize for sorting/pairing
            side = 'a' if kind == 'face' else 'b'
            key = f"{base}__{num_norm}"
            entry = pairs.setdefault(key, {'base': base, 'num': num_norm})
            entry[side] = f
            continue

        m1 = ab_pattern.match(stem)
        if m1:
            base = m1.group(1)
            side = m1.group(2).lower()
            key = base  # legacy key
            entry = pairs.setdefault(key, {'base': base, 'num': None})
            entry[side] = f

    # Build sorted result
    def sort_key(item: Tuple[str, Dict[str, object]]):
        _key, d = item
        base = str(d.get('base', '')).lower()
        num = d.get('num')
        if num is None:
            return (base, -1)
        try:
            return (base, int(str(num)))
        except Exception:
            return (base, 0)

    result: List[Tuple[str, Optional[Path], Optional[Path]]] = []
    for _key, d in sorted(pairs.items(), key=sort_key):
        base = str(d.get('base', ''))
        num = d.get('num')
        base_display = f"{base}[{num}]" if num else base
        result.append((base_display, d.get('a'), d.get('b')))

    return result

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
        name = copyright_name[:COPY_MAX_CHARS]
        c.drawCentredString(page_w / 2.0, y, f"\u00A9 by {name}")

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
        c.line(x, y_top - half, x, y_top + half)
    for y in ys:
        c.line(x_left - half, y, x_left + half, y)
        c.line(x_right - half, y, x_right + half, y)
    c.restoreState()

def draw_corner_marks_grid(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float, cols: int, rows: int):
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)
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
        c.line(x, y_top - half, x, y_top + half)
    for y in ys:
        c.line(x_left - half, y, x_left + half, y)
        c.line(x_right - half, y, x_right + half, y)

    c.restoreState()

def draw_corner_marks_3x3(c: canvas.Canvas, x0: float, y0: float, card_w: float, card_h: float):
    """
    Draw L-shaped corner marks at all 4 corners of the 3x3 grid.
    Each segment is centered on the grid corner, so it is half inside and half outside,
    matching the visual style of draw_outer_marks_3x3().
    """
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_STD)

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

def place_images_grid_inner(c: canvas.Canvas,
                            img_paths: List[Optional[Path]],
                            x0: float, y0: float,
                            card_w: float, card_h: float,
                            cols: int, rows: int,
                            is_back: bool,
                            quality_key: str,
                            card_box_inches: Tuple[float, float]):
    per_page = cols * rows
    for idx, img_path in enumerate(img_paths[:per_page]):
        row = idx // cols
        col = idx % cols
        # Rückseite: Spalten spiegeln
        if is_back:
            col = (cols - 1) - col
        x = x0 + col * card_w
        y = y0 + (rows - 1 - row) * card_h
        if img_path is None or not img_path.exists():
            continue
        processed = preprocess_card_image_for_pdf(img_path, quality_key, card_box_inches)
        draw_w, draw_h = fit_image_into_box(processed, card_w, card_h)
        dx = x + (card_w - draw_w) / 2.0
        dy = y + (card_h - draw_h) / 2.0
        c.drawImage(ImageReader(str(processed)), dx, dy, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")

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
def get_2x3_box_size_pt() -> Tuple[float, float]:
    """
    The actual poker area is still 2.5"x3.5" (180x252 pt),
    but the full image is 825x1125 px, inner is 750x1050 px.
    So the full box is scaled by (825/750, 1125/1050).
    """
    box_w = POKER_W_PT * (BLEED_W_PX / INNER_W_PX)   # 180 * 1.1 = 198
    box_h = POKER_H_PT * (BLEED_H_PX / INNER_H_PX)   # 252 * 1.071428... = 270
    return box_w, box_h

def get_2x3_box_inches() -> Tuple[float, float]:
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
    Works for any cols x rows bleed grid (e.g., 3x2, 5x2).
    """
    c.saveState()
    c.setLineWidth(CUTMARK_LINE_PT_2X3)
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

    L = CUTMARK_LEN_PT_2X3
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

def place_images_gutterfold_grid(c: canvas.Canvas,
                                 pairs_group: List[Tuple[str, Optional[Path], Optional[Path]]],
                                 x0: float, y0: float,
                                 card_w: float, card_h: float,
                                 cols: int,
                                 fold_gutter: float,
                                 quality_key: str,
                                 card_box_inches: Tuple[float, float]):
    per_page = cols
    padded = pairs_group + [("", None, None)] * (per_page - len(pairs_group))
    grid_w = cols * card_w
    grid_h = 2 * card_h + fold_gutter

    y_bottom = y0
    y_top = y0 + card_h + fold_gutter
    fold_y = y0 + card_h + fold_gutter / 2.0

    for col in range(cols):
        _base, front, back = padded[col]
        x = x0 + col * card_w  # bündig

        if front and front.exists():
            processed_f = preprocess_card_image_for_pdf(front, quality_key, card_box_inches)
            draw_image_transformed(c, processed_f, x, y_top, card_w, card_h, rotate_deg=0, mirror_x=False)

        if back and back.exists():
            processed_b = preprocess_card_image_for_pdf(back, quality_key, card_box_inches)
            draw_image_transformed(c, processed_b, x, y_bottom, card_w, card_h, rotate_deg=180, mirror_x=False)

    if GF_DRAW_FOLD_LINE:
        draw_gutterfold_line_horizontal(c, x0, fold_y, grid_w)

    x_marks = [x0 + j * card_w for j in range(cols + 1)]
    y_edges = sorted(set([y0, y0 + card_h, y0 + card_h + fold_gutter, y0 + grid_h]))
    draw_cutmarks_gutterfold(c, x0=x0, y0=y0, grid_w=grid_w, grid_h=grid_h, y_edges=y_edges, x_marks=x_marks)

    y_gutter_bottom = y0 + card_h
    y_gutter_top = y0 + card_h + fold_gutter
    bridge_x = [x0 + j * card_w for j in range(0, cols + 1)]
    draw_gutter_bridge_marks(c, bridge_x, y_gutter_bottom, y_gutter_top)

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
                 include_back_pages: bool = True):
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
                card_box_inches=(POKER_W_PT/72.0, POKER_H_PT/72.0)
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
                    c, backs, x0, y0, card_w, card_h,
                    cols=cols, rows=rows, is_back=True,
                    quality_key=quality_key,
                    card_box_inches=(POKER_W_PT/72.0, POKER_H_PT/72.0)
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
        
        box_w, box_h = get_2x3_box_size_pt()
        # Logo IMMER zeichnen, aber NICHT als harte Reserve
        _apply_logo = bool(logo_path)
        _logo_for_calc = None
        cols, rows, x0, y0, grid_w, grid_h, grid_top_y = compute_max_grid_counts(
            page_w, page_h, box_w, box_h,
            MARGINS_PT, _logo_for_calc, BOTTOM_RESERVED_PT, extra_vertical_pt=0.0
        )
        per_page = cols * rows
        c = create_pdf_canvas(out_path, pagesize_tuple, author=(copyright_name or ''))
        sheet_no = 0
        for group in chunk(pairs, per_page):
            sheet_no += 1
            fronts = [a for (_n, a, _b) in group] + [None] * (per_page - len(group))
            backs  = [b for (_n, _a, b) in group] + [None] * (per_page - len(group))
            place_images_bleed_grid(
                c, fronts, x0, y0, box_w, box_h,
                cols=cols, rows=rows, is_back=False,
                quality_key=quality_key,
                card_box_inches=get_2x3_box_inches()
            )


            if _apply_logo:
                header_h = _compute_header_h_for_logo(logo_path, page_w, page_h, MARGINS_PT, grid_top_y)
                if header_h > 1.0:
                    draw_logo_in_header_band(c, logo_path, page_w, page_h, MARGINS_PT, header_h)
            draw_bottom_line(c, page_w, copyright_name, version_str, f"{sheet_no}a")
            c.showPage()
            if include_back_pages and any(p for p in backs if p and p.exists()):
                place_images_bleed_grid(
                    c, backs, x0, y0, box_w, box_h,
                    cols=cols, rows=rows, is_back=True,
                    quality_key=quality_key,
                   card_box_inches=get_2x3_box_inches()
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
        _apply_logo = bool(logo_path)
        sheet_no = 0
        for group in chunk(pairs, per_page):
            sheet_no += 1
            place_images_gutterfold_grid(
                c, group, x0, y0, card_w, card_h,
                cols=cols, fold_gutter=GF_FOLD_GUTTER_PT,
                quality_key=quality_key,
                card_box_inches=(POKER_W_PT/72.0, POKER_H_PT/72.0)
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

from reportlab.lib.pagesizes import landscape

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
    prompt_language_if_needed()
    clear_tmp_cache()

    print(f'PnP_PDF_Creator {SCRIPT_VERSION} | PIL available: {Image is not None}')
    print("Free software – see LICENSE.txt (and header license notice).")
    print(" ")

    # --- Formatwahl ---
    fmt = prompt_card_format()
    apply_card_format(fmt)
    print_selected_format_info(fmt)
    
    # --- Neue Layout-Auswahl (Standard/Bleed/Gutterfold/All) ---
    layout_keys = prompt_layout_dynamic()


    # --- Papierformate (A4 / Letter / Both) ---
    size_modes = prompt_pagesize_mode()

    # --- Kartenordner ---
    folder = prompt_folder()
    pairs = find_card_pairs(folder)

    if not pairs:
        pause_before_exit(t("no_cards"))
        return

    # =====================================================
    # Shared Cardback Handling
    # =====================================================
    include_back_pages = True
    missing_backs = [base for (base, a, b) in pairs if not (b and b.exists())]
    has_any_back = len(missing_backs) < len(pairs)

    if not has_any_back:
        # Shared fallback
        shared_back = find_named_image_in_folder(folder, CARDBACK_BASENAME)
        if shared_back:
            print(t('using_cardback', file=shared_back.name))
            pairs = [(base, a, shared_back) for (base, a, _b) in pairs]
            missing_backs = []
            has_any_back = True
        else:
            include_back_pages = False
            # Keine Rückseiten → kein Gutterfold möglich
            if any(k.lower() == "gutterfold" for k in layout_keys):
                layout_keys = [k for k in layout_keys if k.lower() != "gutterfold"]
                print(t('skip_gutterfold_no_backs', name=CARDBACK_BASENAME))

    # Teilweise fehlende Backs → Gutterfold entfernen
    if missing_backs and any(k.lower() == "gutterfold" for k in layout_keys):
        # fronts ohne backs anzeigen
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

    # =====================================================
    # Bildanalyse (Größen, Bleed-tauglichkeit)
    # =====================================================
    sizes, too_small, too_small_bleed, pairs_for_2x3, skipped_2x3 = analyze_card_images(pairs)

    # Hinweis: kleine Karten → skalieren
    if too_small:
        names = sorted({p.name for p in too_small})
        print(t('warn_too_small_upscale',
                minw=INNER_W_PX, minh=INNER_H_PX,
                count=len(names), files=', '.join(names)))

    # Spezifisch: Wenn irgendwo Bleed-Mindestgröße unterschritten wird,
    # sollen alle Bleed-Layouts (bleed, 2x3, 2x5) NICHT erzeugt werden.
    requested_bleed_any = any(k.lower() in ("bleed", "2x3", "2x5") for k in layout_keys)
    if requested_bleed_any and too_small_bleed:
        names = sorted({p.name for p in too_small_bleed})
        if len(names) > 30:
            shown = ', '.join(names[:30]) + f" ... (+{len(names)-30})"
        else:
            shown = ', '.join(names)
        # Entferne alle Bleed-Varianten
        layout_keys = [k for k in layout_keys if k.lower() not in ("bleed", "2x3", "2x5")]
        print(t("skip_bleed_due_to_small",
                minw=BLEED_W_PX, minh=BLEED_H_PX,
                count=len(names), files=shown))

    # Prüfen, ob 2x3 oder 2x5 überhaupt möglich sind (Bleed nötig), falls Bleed nicht bereits global deaktiviert wurde
    requested_2x3 = any(k.lower() == "2x3" for k in layout_keys)
    requested_2x5 = any(k.lower() == "2x5" for k in layout_keys)

    if requested_2x3 or requested_2x5:
        if not pairs_for_2x3 or len(pairs_for_2x3) != len(pairs):
            if requested_2x3:
                layout_keys = [k for k in layout_keys if k.lower() != "2x3"]
                print(t('skip_2x3', minw=BLEED_W_PX, minh=BLEED_H_PX))

            if requested_2x5:
                layout_keys = [k for k in layout_keys if k.lower() != "2x5"]
                print(t('skip_2x5', minw=BLEED_W_PX, minh=BLEED_H_PX))

    if not layout_keys:
        pause_before_exit()
        return

    # =====================================================
    # Weitere Eingaben
    # =====================================================
    logo_path = prompt_logo_path(folder)
    quality_key = prompt_quality()
    copyright_name = prompt_copyright_name()
    version_str = prompt_version()
    out_base = prompt_output_base("cards")

    multi = len(size_modes) > 1

    # =====================================================
    # PDF-Erzeugung
    # =====================================================
    for layout_key in layout_keys:
        # Paare je nach Modus (Bleed braucht Bleed-taugliche Paare)
        if layout_key in ("bleed", "2x3", "2x5"):
            pairs_layout = pairs_for_2x3
        else:
            pairs_layout = pairs
        for base_pagesize, suffix in size_modes:
            # Orientierung je Layout
            if layout_key in ("gutterfold",):
                pagesize_tuple = choose_gutterfold_orientation(base_pagesize)
            elif layout_key in ("bleed", "2x3", "2x5"):
                pagesize_tuple = landscape(base_pagesize)
            else:
                pagesize_tuple = base_pagesize
                
            # Dateiname (Suffix)
            if layout_key in ("standard", "3x3", "3x4"):
                layout_suffix = "_standard"
            elif layout_key in ("bleed", "2x3", "2x5"):
                layout_suffix = "_bleed"
            else:
                layout_suffix = "_gutterfold"
            out_path = (Path(f"{out_base}{layout_suffix}{suffix}.pdf").resolve()
                        if multi else Path(f"{out_base}{layout_suffix}.pdf").resolve())
            generate_pdf(
                layout_key=layout_key,
                out_path=out_path,
                pagesize_tuple=pagesize_tuple,
                pairs=pairs_layout,
                logo_path=logo_path,
                copyright_name=copyright_name,
                version_str=version_str,
                quality_key=quality_key
            )
            print(t("done", path=out_path))

if __name__ == "__main__":
    main()
