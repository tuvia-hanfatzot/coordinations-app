import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from docxtpl import DocxTemplate
import pyproj
from docx import Document
from docx.shared import Pt
from docx.shared import RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
import re
import matplotlib
matplotlib.use("Agg")  # headless backend for Streamlit/servers
import matplotlib.pyplot as plt
import math
from docxtpl import InlineImage
from docx.shared import Cm
import contextily as cx
from pyproj import Transformer

# Put this near the top (after imports)
transformer = pyproj.Transformer.from_crs("EPSG:4326", "EPSG:32636", always_xy=True)

# Constants
TEMPLATE_FILE = "coordinations_template.docx"
HEBREW_LETTERS = ["א","ב","ג","ד","ה","ו","ז","ח","ט","י","יא","יב","יג","יד","טו","טז","יז","יח","יט","כ","כא","כב","כג","כד","כה","כו","כז","כח","כט","ל"]

def hebrew_index(idx_zero_based: int) -> str:
    """Return a Hebrew index label; falls back to 1-based numerals if we run out."""
    if idx_zero_based < len(HEBREW_LETTERS):
        return HEBREW_LETTERS[idx_zero_based]
    return str(idx_zero_based + 1)

PERSONNEL_TRANSLATIONS = {
    "driver": "נהג",
    "assistant driver": "נהג משנה",
    "passenger": "נוסע",
    "patient": "מטופל",
    "patient's companion": "מלווה",
    "security escort": "נוסע - מאבטח",
    "suno": "בכיר",
    "mission team leader": "נוסע",
    "south to north": "\u200F)מדרום לצפון(\u200F",
    "north to south": "\u200F)מצפון לדרום(\u200F",
    "leaving gaza": "\u200F)יוצא מרצע(\u200F",
    "entering gaza": "\u200F)נכנס לרצע(\u200F"
}

# SUNO / VIP titles (normalized as UPPER, no spaces)
SUNO_TITLES = {
    "SUNC80482D": "ר' OCHA ישראל",
    "SUNJ14406D": "ר' OCHA ישראל",
    "SUNJ16507": "סגן ר' OCHA ישראל",
    "SUNB28399": "סגן ר' OCHA עזה",
    "SUNJ17294": "יועצת מיוחדת של OCHA",
    "SUNR32190": "סגנית המתאמת ההומניטרית",
    "SUNR23865D": "מ\"מ המתאם ההומניטרי",
    "AUNR72415": "ר' WHO ישראל",
    "AUNR73490D": "ר' WFP ישראל",
    "AUNB54993D": "ר' WFP עזה",
    "124500670": "סגן ר' WFP עזה",
    "SUNR33194D": "ר' UNICEF ישראל",
    "SUNJ10442D": "ר' UNICEF עזה",
    "144666595": "ר' UNMAS עזה",
    "SUNC80122D": "ר' UNDP עזה",
    "SUNJ10610D": "סגן ר' UNDP ישראל",
    "SUNC80167D": "ר' UNDSS ישראל",
    "SUNR33736D": "סגן ר' UNDSS ישראל",
    "X0N73G02": "ר' צלא עזה",
    "PB3553006": "סגן ר' צלא ישראל",
    "567985289": "מייסד ארגון WCK",
    "548965410": "ר' ANERA עזה",
}

def _norm_id_for_lookup(pid: str) -> str:
    return re.sub(r"\s+", "", str(pid or "")).upper()

def _get_suno_title(pid: str):
    return SUNO_TITLES.get(_norm_id_for_lookup(pid))

VEHICLE_TRANSLATIONS = {
    "truck": "משאית",
    "box truck": "משאית סגורה",
    "flatbed": "משאית פתוחה",
    "bus": "אוטובוס",
    "un": "טויוטה לנד קרוזר", 
    "lc": "טויטה לנד קרוזר",
    "toyota land cruiser": "טויוטה לנד קרוזר",
    "land cruiser": "טויוטה לנד קרוזר"  
}

# === Site name translations (lowercased keys; substring match) ===
SITE_NAME_TRANSLATIONS = {
    "kerem shalom": "כרמש",
    "kerem shaloum": "כרמש",
    "karem salim": "כרמש",
    "karem salem": "כרמש",
    "ks": "כרמש",
    "kissufim": "מעבר 147",
    "kissuffim": "מעבר 147",
    "kisuffim": "מעבר 147",
    "ksf": "מעבר 147",
    "gate 96": "שער 96",
    "airport road": "פולאט",
}

def translate_site_name(name: str) -> str:
    s = (name or "").strip()
    low = s.lower()
    for key, heb in SITE_NAME_TRANSLATIONS.items():
        if key in low:
            return heb
    return s

# Utility Functions
def extract_below_target(df, target, max_offset=3):
    """
    Extract the value below a header.
    - 'Organization' → exact match only
    - 'Date of coordination' (or 'Date') → prefer any cell containing 'date of coordination';
      else whole-cell 'Date' (avoids 'Coordinates')
    - Others → exact-first, then substring
    """
    tgt = target.strip().lower()
    nrows, ncols = len(df), len(df.columns)

    def _scan_from(r, c):
        for off in range(1, max_offset + 1):
            nr = r + off
            if nr >= nrows: break
            nxt = df.iloc[nr, c]
            if isinstance(nxt, str) and nxt.strip():
                return nxt.strip()
            elif not pd.isna(nxt):
                return str(nxt)
        return "[MISSING DATA BELOW TARGET]"

    # --- Organization: exact only ---
    if tgt == "organization":
        for r in range(nrows - max_offset):
            for c in range(ncols):
                if str(df.iloc[r, c]).strip().lower() == "organization":
                    return _scan_from(r, c)
        return "[TARGET NOT FOUND]"

    # --- Dates: prefer 'Date of coordination' over plain 'Date' ---
    if tgt in ("date", "date of coordination"):
        # 1) Any cell that CONTAINS 'date of coordination'
        for r in range(nrows - max_offset):
            for c in range(ncols):
                if "date of coordination" in str(df.iloc[r, c]).strip().lower():
                    return _scan_from(r, c)
        # 2) Whole-cell 'Date' (avoids matching 'Coordinates')
        wb = re.compile(r"^\s*date\s*$", re.IGNORECASE)
        for r in range(nrows - max_offset):
            for c in range(ncols):
                if wb.match(str(df.iloc[r, c])):
                    return _scan_from(r, c)
        return "[TARGET NOT FOUND]"

    # --- Generic: exact-first, then substring ---
    for r in range(nrows - max_offset):
        for c in range(ncols):
            if str(df.iloc[r, c]).strip().lower() == tgt:
                return _scan_from(r, c)

    for r in range(nrows - max_offset):
        for c in range(ncols):
            if tgt in str(df.iloc[r, c]).strip().lower():
                return _scan_from(r, c)

    return "[TARGET NOT FOUND]"

def extract_hours(df):
    """
    Start time  = first valid time under "Time (hh:mm) of requested departure".
    End time    = last  valid time under "Time (hh:mm) of expected arrival".
    Blanks are ignored; we scan to the bottom of the sheet.
    Arrival is then shifted by +5h (keep or change OFFSET_HOURS as needed).
    """
    OFFSET_HOURS = 5  # change if needed

    def _norm_hhmm(s: str) -> str:
        m = re.match(r'^\s*([0-2]?\d):([0-5]\d)\s*$', s)
        if m:
            return f"{int(m.group(1)):02d}:{m.group(2)}"
        return s

    def _coerce_time(value):
        # Accept HH:MM, HH:MM:SS, Excel serials, or text containing a time.
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        try:
            if isinstance(value, (int, float)):
                base = datetime(1899, 12, 30) + timedelta(days=float(value))
                return base.strftime("%H:%M")
            s = str(value).strip()

            # direct formats
            for fmt in ("%H:%M", "%H:%M:%S"):
                try:
                    return datetime.strptime(s, fmt).strftime("%H:%M")
                except ValueError:
                    pass

            # time embedded in text
            m = re.search(r'([01]?\d|2[0-3]):[0-5]\d', s)
            if m:
                return _norm_hhmm(m.group(0))

            # numeric-as-string → Excel serial
            f = float(s)
            base = datetime(1899, 12, 30) + timedelta(days=f)
            return base.strftime("%H:%M")
        except Exception:
            return None

    def _find_header(df, labels):
        labels = [l.lower() for l in labels]
        for r in range(len(df)):
            for c in range(len(df.columns)):
                txt = str(df.iloc[r, c]).strip().lower()
                if any(l in txt for l in labels):
                    return r, c
        return None, None

    def _scan_column_for_times(df, start_row, col, want="first"):
        if start_row is None or col is None:
            return "[MISSING TIME]"
        times = []
        for r in range(start_row + 1, len(df)):  # scan to the bottom; no early stopping
            t = _coerce_time(df.iloc[r, col])
            if t:
                times.append(_norm_hhmm(t))
        if not times:
            return "[MISSING TIME]"
        return times[0] if want == "first" else times[-1]

    # Find the two headers (robust to phrasing)
    dep_row, dep_col = _find_header(df, ["time (hh:mm) of requested departure", "requested departure"])
    arr_row, arr_col = _find_header(df, ["time (hh:mm) of expected arrival", "expected arrival"])

    start_time = _scan_column_for_times(df, dep_row, dep_col, want="first")
    end_time   = _scan_column_for_times(df, arr_row, arr_col, want="last")

    # Apply +5h to arrival if found
    if end_time != "[MISSING TIME]":
        end_dt = datetime.strptime(end_time, "%H:%M") + timedelta(hours=OFFSET_HOURS)
        end_time = end_dt.strftime("%H:%M")

    return start_time, end_time

def format_and_validate_date(date_value):
    """
    Returns date as dd.MM.yyyy.
    Accepts: Excel serials; datetime; 'dd.MM.yyyy'; 'dd/MM/yyyy';
             with optional time, e.g. '17/09/2025 12:00', '17.09.2025 06:30'.
    """
    if isinstance(date_value, (int, float)):
        try:
            d = datetime(1899, 12, 30) + timedelta(days=float(date_value))
            return d.strftime("%d.%m.%Y")
        except Exception:
            return "[INVALID DATE]"

    if isinstance(date_value, datetime):
        return date_value.strftime("%d.%m.%Y")

    if isinstance(date_value, str):
        s = date_value.strip()
        # Normalise whitespace
        s = re.sub(r"\s+", " ", s)

        # Try explicit formats first
        fmts = [
            "%d.%m.%Y", "%d/%m/%Y",
            "%d.%m.%Y %H:%M", "%d/%m/%Y %H:%M",
            "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"
        ]
        for fmt in fmts:
            try:
                d = datetime.strptime(s, fmt)
                return d.strftime("%d.%m.%Y")
            except ValueError:
                pass

        # Last resort: regex extract dd/mm/yyyy or dd.mm.yyyy at start of string
        m = re.search(r"\b(\d{1,2})[./](\d{1,2})[./](\d{4})\b", s)
        if m:
            try:
                d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                return d.strftime("%d.%m.%Y")
            except ValueError:
                return "[INVALID DATE]"

    return "[INVALID DATE]"

def latlon_to_utm(lat, lon):
    """Converts latitude/longitude to UTM and shortens the northing value."""
    try:
        easting, northing = transformer.transform(lon, lat)
        short_northing = str(int(northing))[1:]  # Remove the leading '3'
        return f"{int(easting)}/{short_northing}"
    except Exception:
        return "[INVALID COORDS]"

def _is_site_name_header(text: str) -> bool:
    t = text.lower()
    return "site name" in t or ("site" in t and "name" in t)

def _is_coordinates_header(text: str) -> bool:
    t = text.lower()
    return ("coordinates" in t or "coordinate" in t or "coords" in t or "gps" in t)

def _parse_lat_lon(coord_text: str):
    """
    Accepts:
      - Decimal:   31.533070, 34.456701   (comma/space/semicolon separators; ° optional)
      - DMS:       31°25'39.5"N 34°23'10.4"E   (minutes/seconds marks and N/S/E/W)
    Returns (lat, lon) as floats, or None if invalid/out-of-range.
    """
    s = str(coord_text or "").strip()

    # --- 1) Decimal degrees ---
    m = re.match(r'^\s*([+-]?\d{1,2}\.\d+)\s*[°º]?\s*[,;\s]\s*([+-]?\d{1,3}\.\d+)\s*[°º]?\s*$',
                 s)
    if m:
        lat = float(m.group(1)); lon = float(m.group(2))
        if -90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0:
            return lat, lon
        return None

    # --- 2) DMS (degrees° minutes' seconds" HEMI) ---
    # seconds optional, quotes/marks flexible, allows spaces
    dms_re = re.compile(
        r'([+-]?\d{1,3})\s*[°º]\s*'          # degrees
        r'(\d{1,2})\s*[\']?\s*'              # minutes
        r'(?:(\d{1,2}(?:\.\d+)?))?\s*(?:["”″])?\s*'  # seconds (optional)
        r'([NSEW])',                          # hemisphere
        re.IGNORECASE
    )
    parts = list(dms_re.finditer(s))
    if len(parts) >= 2:
        def dms_to_dd(deg_s, min_s, sec_s, hemi):
            deg = float(deg_s)
            mins = float(min_s)
            secs = float(sec_s) if sec_s is not None else 0.0
            dd = abs(deg) + mins/60.0 + secs/3600.0
            if (str(hemi).upper() in ("S", "W")) or (str(deg_s).startswith("-")):
                dd = -dd
            return dd

        # Assign by hemisphere
        lat_dd = lon_dd = None
        for m in parts[:2]:
            deg_s, min_s, sec_s, hemi = m.groups()
            dd = dms_to_dd(deg_s, min_s, sec_s, hemi)
            if hemi.upper() in ("N", "S"):
                lat_dd = dd
            else:
                lon_dd = dd

        if lat_dd is not None and lon_dd is not None:
            if -90.0 <= lat_dd <= 90.0 and -180.0 <= lon_dd <= 180.0:
                return lat_dd, lon_dd

    return None

def count_coordinates_levels(df):
    """Extract unique (ordered) site rows; stop at other sections.
    Include rows that have a name but missing/invalid coordinates (flagged)."""
    location_col, coordinates_col, start_row = None, None, None

    # Find headers
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip()
            if _is_site_name_header(cell_value):
                location_col = col; start_row = row + 1
            elif _is_coordinates_header(cell_value):
                coordinates_col = col; start_row = row + 1

    if location_col is None or coordinates_col is None or start_row is None:
        return ["[MISSING LOCATION/COORDINATES]"]

    extracted_levels, last_key = [], None

    for r in range(start_row, len(df)):
        if _row_has_end_marker(df, r):
            break

        name  = df.iloc[r, location_col]
        coords = df.iloc[r, coordinates_col]

        # Clean + translate site name
        clean_name = translate_site_name(_clean_site_name(name)) if isinstance(name, str) else ""

        # If there's no usable name, skip the row entirely
        if not clean_name:
            continue

        parsed = _parse_lat_lon(coords)

        if parsed:
            lat, lon = parsed
            key = (round(lat, 6), round(lon, 6))
            if last_key is not None and key == last_key:
                continue  # skip consecutive duplicates
            utm_coords = latlon_to_utm(lat, lon)
            last_key = key
        else:
            # No/invalid coordinates -> include the row and flag it
            utm_coords = "[MISSING LOCATION/COORDINATES]"

        idx = len(extracted_levels)
        letter = hebrew_index(idx)
        extracted_levels.append((letter, clean_name, utm_coords))

    if not extracted_levels:
        return ["[MISSING LOCATION/COORDINATES]"]

    lines = []
    for i, (letter, location, utm) in enumerate(extracted_levels):
        if i == 0:
            point_type = "נקודת התחלה"
        elif i == len(extracted_levels) - 1:
            point_type = "נקודת סיום"
        else:
            point_type = "נקודת עצירה"
        lines.append(f"\u202B{letter}.\u00A0{point_type} -\u00A0\u200F{location}\u200F \u200E{utm}\u200E")

    return lines

END_SECTION_TOKENS = [
    "first name", "last name", "mission of the specific personnel",
    "vehicle type", "plate number", "equipment", "personnel"
]

def _row_has_end_marker(df, row) -> bool:
    # If any cell in this row contains a token that signals the next section, stop.
    row_texts = [str(x).strip().lower() for x in df.iloc[row].tolist()]
    return any(any(tok in cell for tok in END_SECTION_TOKENS) for cell in row_texts)

def _clean_site_name(name: str) -> str:
    s = (name or "").strip()
    return re.sub(r'^\s*(from|to)\s*:\s*', '', s, flags=re.IGNORECASE)

def _parse_lat_lon(coord_text: str):
    """
    Accepts:
      - Decimal:   31.533070, 34.456701   (comma/space/semicolon separators; ° optional)
      - DMS:       31°25'39.5"N 34°23'10.4"E   (minutes/seconds marks and N/S/E/W)
    Returns (lat, lon) as floats, or None if invalid/out-of-range.
    """
    s = str(coord_text or "").strip()

    # --- 1) Decimal degrees ---
    m = re.match(r'^\s*([+-]?\d{1,2}\.\d+)\s*[°º]?\s*[,;\s]\s*([+-]?\d{1,3}\.\d+)\s*[°º]?\s*$',
                 s)
    if m:
        lat = float(m.group(1)); lon = float(m.group(2))
        if -90.0 <= lat <= 90.0 and -180.0 <= lon <= 180.0:
            return lat, lon
        return None

    # --- 2) DMS (degrees° minutes' seconds" HEMI) ---
    # seconds optional, quotes/marks flexible, allows spaces
    dms_re = re.compile(
        r'([+-]?\d{1,3})\s*[°º]\s*'          # degrees
        r'(\d{1,2})\s*[\']?\s*'              # minutes
        r'(?:(\d{1,2}(?:\.\d+)?))?\s*(?:["”″])?\s*'  # seconds (optional)
        r'([NSEW])',                          # hemisphere
        re.IGNORECASE
    )
    parts = list(dms_re.finditer(s))
    if len(parts) >= 2:
        def dms_to_dd(deg_s, min_s, sec_s, hemi):
            deg = float(deg_s)
            mins = float(min_s)
            secs = float(sec_s) if sec_s is not None else 0.0
            dd = abs(deg) + mins/60.0 + secs/3600.0
            if (str(hemi).upper() in ("S", "W")) or (str(deg_s).startswith("-")):
                dd = -dd
            return dd

        # Assign by hemisphere
        lat_dd = lon_dd = None
        for m in parts[:2]:
            deg_s, min_s, sec_s, hemi = m.groups()
            dd = dms_to_dd(deg_s, min_s, sec_s, hemi)
            if hemi.upper() in ("N", "S"):
                lat_dd = dd
            else:
                lon_dd = dd

        if lat_dd is not None and lon_dd is not None:
            if -90.0 <= lat_dd <= 90.0 and -180.0 <= lon_dd <= 180.0:
                return lat_dd, lon_dd

    return None

def extract_sites_for_map(df):
    """Return ordered, de-duplicated valid sites for mapping."""
    location_col, coordinates_col, start_row = None, None, None

    # Find headers
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip()
            if _is_site_name_header(cell_value):
                location_col = col; start_row = row + 1
            elif _is_coordinates_header(cell_value):
                coordinates_col = col; start_row = row + 1

    if location_col is None or coordinates_col is None or start_row is None:
        return []

    sites, last_key = [], None

    for r in range(start_row, len(df)):
        if _row_has_end_marker(df, r):
            break

        name  = df.iloc[r, location_col]
        coords = df.iloc[r, coordinates_col]

        parsed = _parse_lat_lon(coords)
        clean_name = _clean_site_name(name) if isinstance(name, str) else ""

        if not (parsed and clean_name):
            continue

        lat, lon = parsed
        key = (round(lat, 6), round(lon, 6))
        if last_key is not None and key == last_key:
            continue

        sites.append({"name": clean_name, "lat": lat, "lon": lon})
        last_key = key

    return sites

def generate_sites_map_image(sites, out_path="sites_map.png"):
    """
    Renders an OSM basemap with labelled points.
    Robust to empty/invalid points and single-point extents.
    """
    if not sites:
        return None

    wm = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)

    xs, ys, kept = [], [], []
    for s in sites:
        try:
            x, y = wm.transform(s["lon"], s["lat"])
            if math.isfinite(x) and math.isfinite(y):
                xs.append(x); ys.append(y); kept.append(s)
        except Exception:
            continue

    if not xs:  # nothing valid
        return None

    fig, ax = plt.subplots(figsize=(6.5, 5), dpi=200)
    ax.scatter(xs, ys, s=30, zorder=3)

    for (x, y), s in zip(zip(xs, ys), kept):
        ax.annotate(
            f"{s['name']}\n{s['lat']:.6f}, {s['lon']:.6f}",
            (x, y), xytext=(0, 10), textcoords="offset points",
            ha="center", va="bottom", fontsize=9, zorder=4,
            bbox=dict(facecolor="white", alpha=0.7, edgecolor="none", pad=1.5)
        )

    # Safe viewport
    if len(xs) == 1:
        pad_x = pad_y = 500  # meters around the single point
        ax.set_xlim(xs[0] - pad_x, xs[0] + pad_x)
        ax.set_ylim(ys[0] - pad_y, ys[0] + pad_y)
    else:
        pad_x = max(200, (max(xs) - min(xs)) * 0.15)
        pad_y = max(200, (max(ys) - min(ys)) * 0.15)
        ax.set_xlim(min(xs) - pad_x, max(xs) + pad_x)
        ax.set_ylim(min(ys) - pad_y, max(ys) + pad_y)

    # Basemap (guarded; still renders points if tiles fail/offline)
    try:
        cx.add_basemap(ax, source=cx.providers.OpenStreetMap.Mapnik, crs="EPSG:3857", attribution=False)
        ax.set_axis_off()
    except Exception:
        ax.grid(True, linestyle=":", linewidth=0.5)

    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches="tight", pad_inches=0.05)
    plt.close(fig)
    return out_path

def check_white_wheels(ids):
    """Ensures white wheels check is independent of other data."""
    for id_value in ids:
        if isinstance(id_value, str) and id_value.strip():
            first_digit = id_value.strip()[0]
            if first_digit not in {"3", "4", "8", "9"}:
                return "גלגלים לבנים"
    return ""

def translate_personnel_role(text):
    if not isinstance(text, str) or not text.strip():
        return "[MISSING DATA]"  # IMPORTANT: Capital letters

    original_text = text.strip()  # Save original casing
    text = original_text.lower()  # Use lowercase only for matching
    main_role = ""
    notes = []

    for eng_term, heb_term in PERSONNEL_TRANSLATIONS.items():
        if eng_term in text:
            if "(" in heb_term:  # Direction/note
                notes.append(heb_term)
            else:  # Main role
                main_role = heb_term

    if main_role:
        if notes:
            return f"\u202B{main_role} {' '.join(notes)}\u202C"
        else:
            return f"\u202B{main_role}\u202C"
    else:
        if notes:
            return f"\u202B{' '.join(notes)}\u202C"
        else:
            return original_text  # keep as is if nothing found

def _combine_full_name(first_name: str, last_name: str) -> str:
    fn = (first_name or "").strip()
    ln = (last_name or "").strip()

    if not fn and not ln:
        return "[MISSING DATA]"

    # If both cells are identical (e.g., both already contain "John Doe")
    if fn and ln and fn.lower() == ln.lower():
        return fn

    # If one cell already contains the other (handles cases where a column has the full name)
    if fn and ln:
        if fn.lower() in ln.lower() and len(ln) >= len(fn):
            return ln
        if ln.lower() in fn.lower() and len(fn) >= len(ln):
            return fn

        # Otherwise, concatenate but dedupe repeating tokens: "John John Doe" -> "John Doe"
        parts = (fn + " " + ln).split()
        seen = set()
        deduped = []
        for p in parts:
            key = p.lower()
            if key not in seen:
                seen.add(key)
                deduped.append(p)
        return " ".join(deduped)

    # Only one side present
    return fn or ln

def _clean_id_text(s: str) -> str:
    """
    Extracts ID token: supports 'ID: 900216987', 'ID: AUNB44657', 'id- 12345'.
    Returns the captured token or the original stripped text if no match.
    """
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = s.strip()
    m = re.search(r"\bID\b\s*[:\-]?\s*([A-Z0-9]+)\b", s, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    return s

def extract_personnel_table_data(df):
    """
    Extracts personnel table data for the new format:
    - First Name
    - Last Name
    - ID
    - Mission of the specific personnel  (used as 'notes' / role text)
    Returns (extracted_rows, white_wheels_text).
    """
    # Headers to find (lowercased, matched by substring anywhere in a candidate header row)
    required_headers = {
        "first name": None,
        "last name": None,
        "id": None,
        "mission of the specific personnel": None,
    }

    personnel_start_row = None

    # Find a row that contains all required headers
    for row in range(len(df)):
        row_values = [str(cell).strip().lower() for cell in df.iloc[row]]
        if all(any(key in cell for cell in row_values) for key in required_headers.keys()):
            personnel_start_row = row
            # Map each header to its column index (first occurrence)
            for col in range(len(df.columns)):
                cell_value = str(df.iloc[row, col]).strip().lower()
                for key in list(required_headers.keys()):
                    if key in cell_value and required_headers[key] is None:
                        required_headers[key] = col
            break

    # If not all columns found
    if personnel_start_row is None or any(v is None for v in required_headers.values()):
        return [], "[MISSING PERSONNEL DATA]"

    extracted_data = []
    id_list = []
    empty_row_count = 0
    max_empty_rows_allowed = 1

    # Read rows under the header
    for row in range(personnel_start_row + 1, len(df)):
        # Pull raw cells
        def get_cell(col_idx):
            if col_idx is None:
                return ""
            val = df.iloc[row, col_idx]
            if pd.isna(val) or str(val).strip() == "":
                return ""
            if isinstance(val, float):
                # Excel numeric IDs etc.
                return str(int(val))
            return str(val).strip()

        first_name = get_cell(required_headers["first name"])
        last_name = get_cell(required_headers["last name"])
        pid = get_cell(required_headers["id"])
        pid = _clean_id_text(pid)
        mission_txt = get_cell(required_headers["mission of the specific personnel"])

        # Determine emptiness
        is_empty_row = not any([first_name, last_name, pid, mission_txt])

        if is_empty_row:
            empty_row_count += 1
            if empty_row_count > max_empty_rows_allowed:
                break
            else:
                continue
        else:
            empty_row_count = 0


        full_name = _combine_full_name(first_name, last_name)

        row_data = {
            "name": full_name,
            "id": pid if pid else "[MISSING DATA]",
            # Keep key 'notes' so downstream functions continue to work
            "notes": mission_txt if mission_txt else "[MISSING DATA]",
        }

        if pid:
            id_list.append(pid)

        extracted_data.append(row_data)

    white_wheels_text = check_white_wheels(id_list)
    return extracted_data, white_wheels_text

def update_personnel_table(doc_path, extracted_data, white_wheels_text):
    """Updates the personnel table in the Word document with extracted Excel data."""
    doc = Document(doc_path)

    if not doc.tables:
        return "[NO TABLE FOUND IN DOCUMENT]"

    table = doc.tables[0]  # Assuming first table is the personnel table

    # Clear existing rows except for the header
    while len(table.rows) > 1:
        table._element.remove(table.rows[1]._element)

    # Populate the table with extracted data
    for index, row_data in enumerate(extracted_data, start=1):
        new_row = table.add_row().cells

        number_txt = f".{index}"                                    
        role_txt   = translate_personnel_role(row_data.get("notes", ""))   
        name_txt   = row_data.get("name", "[MISSING DATA]")         
        id_txt     = row_data.get("id", "[MISSING DATA]")           

        suno_title = _get_suno_title(id_txt)                        
        if suno_title:                                              
            name_txt = f"{name_txt} - {suno_title}"                

        # (kept behaviour) assign texts to cells
        new_row[0].text = number_txt            # Numbering (was: f".{index}")
        new_row[1].text = role_txt              # Role (was: translate_personnel_role(...))
        new_row[2].text = name_txt              # Name (was: row_data.get("name", ...))
        new_row[3].text = id_txt                # ID   (was: row_data.get("id", ...))

        for cell in new_row:
            set_font(cell)

        if suno_title:                                               
            for ci in (2, 3):                                        
                for p in new_row[ci].paragraphs:                     
                    for r in p.runs:                                 
                        r.font.bold = True                           

    # Handle the "גלגלים לבנים" placeholder and remove text box if empty
    if white_wheels_text.strip():
        for para in doc.paragraphs:
            if "{ white_wheels }" in para.text:
                para.text = para.text.replace("{ white_wheels }", white_wheels_text)
    else:
        for shape in doc.element.xpath("//w:drawing"):  # Locate all drawing elements (text boxes)
            parent = shape.getparent()
            if parent is not None:
                parent.remove(shape)  # Remove the entire text box
                break

    doc.save("updated_document.docx")

    return "updated_document.docx"

def translate_vehicle_type(vehicle_type):
    """Ensures Hebrew appears first with a visible space before the English text, handling RTL issues in Word."""
    if not isinstance(vehicle_type, str) or not vehicle_type.strip():
        return "[MISSING DATA]"

    normalized_type = vehicle_type.lower().strip()

    for eng_type, hebrew_translation in VEHICLE_TRANSLATIONS.items():
        if eng_type in normalized_type:
            brand_name = normalized_type.replace(eng_type, "").strip().title()
            if brand_name:
                return f"{brand_name}\u00A0{hebrew_translation}"  # Hebrew first, non-breaking space, then English
            else:
                return f"{hebrew_translation}"  # No brand name, just Hebrew

    return vehicle_type  # If no match, return original input

def extract_vehicle_table_data(df):
    """
    Extracts vehicle type and plate number properly, ensuring no empty rows are added,
    while allowing rows to appear after a blank one.
    """
    target_headers = {"vehicle type": None, "plate number": None}
    vehicle_start_row = None
    empty_row_count = 0
    max_empty_rows_allowed = 1

    # Identify the header row dynamically (substring match, not exact)
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            for key in list(target_headers.keys()):
                if key in cell_value and target_headers[key] is None:
                    target_headers[key] = col
                    if vehicle_start_row is None:
                        vehicle_start_row = row

        # If we found both columns on this row, we can stop scanning further rows
        if all(v is not None for v in target_headers.values()):
            break

    if any(v is None for v in target_headers.values()) or vehicle_start_row is None:
        # Keep the existing behaviour so the caller can handle it
        return [], "[MISSING VEHICLE DATA]"

    extracted_vehicles = []

    for row in range(vehicle_start_row + 1, len(df)):
        is_empty_row = True
        row_data = {}

        for key, col_index in target_headers.items():
            val = df.iloc[row, col_index]
            if pd.isna(val) or str(val).strip() == "":
                cell_value = "[MISSING DATA]"
            elif isinstance(val, float):
                cell_value = str(int(val))
            else:
                cell_value = str(val).strip()

            if cell_value != "[MISSING DATA]":
                is_empty_row = False

            row_data[key] = cell_value

        if is_empty_row:
            empty_row_count += 1
            if empty_row_count > max_empty_rows_allowed:
                break
        else:
            empty_row_count = 0
            extracted_vehicles.append(row_data)

    return extracted_vehicles

def update_vehicle_table(doc_path, extracted_vehicles):
    """Updates the vehicle table in the Word document correctly."""
    doc = Document(doc_path)

    if len(doc.tables) < 2:
        return "[SECOND TABLE NOT FOUND]"

    table = doc.tables[1]  # Second table assumed to be vehicle data

    # Clear existing rows except for the header
    while len(table.rows) > 1:
        table._element.remove(table.rows[1]._element)

    # Populate the second table with vehicle data
    for index, row_data in enumerate(extracted_vehicles, start=1):
        new_row = table.add_row().cells
        new_row[0].text = f".{index}"  # Numbering
        new_row[1].text = translate_vehicle_type(row_data.get("vehicle type", "[MISSING DATA]"))  # Translated type
        new_row[2].text = row_data.get("plate number", "[MISSING DATA]")  # Plate number

        for cell in new_row:
            set_font(cell)

    doc.save("updated_document.docx")

    return "updated_document.docx"

def extract_equipment_status(df):
    """
    Scans rows between 'Equipment' and two rows above 'NOTES' to determine if any cell contains 'yes'.
    Returns 'יש ציוד' if found, else 'אין ציוד'.
    """
    equipment_row, equipment_col = None, None
    notes_row = None

    # First: find 'Equipment' position
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if "equipment" in cell_value:
                equipment_row = row
                equipment_col = col
                break
        if equipment_row is not None:
            break

    # Second: find 'NOTES' position
    for row in range(len(df)):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if "mission of the specific personnel" in cell_value:
                notes_row = row
                break
        if notes_row is not None:
            break

    # Validate we found both
    if equipment_row is None or notes_row is None:
        return "[MISSING EQUIPMENT OR NOTES]"

    # Scan between equipment_row+1 and notes_row-2
    for row in range(equipment_row + 1, notes_row - 1):
        for col in range(len(df.columns)):
            cell_value = str(df.iloc[row, col]).strip().lower()
            if cell_value == "yes":
                return "יש ציוד"

    # If no 'yes' found
    return "אין ציוד"

# Function to set font properties in Word tables
def set_font(cell, font_name="David", font_size=12, align_center=True, vcenter=True):
    if vcenter:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for paragraph in cell.paragraphs:
        if align_center:
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)

def highlight_text(doc, text, highlight_color):
    """Highlights text in the document with the given color."""
    for para in doc.paragraphs:
        if text in para.text:
            for run in para.runs:
                if text in run.text:
                    run.font.highlight_color = highlight_color

def update_highlighting(doc_path, start_date, equipment_list):
    from docx.enum.text import WD_COLOR_INDEX
    doc = Document(doc_path)

    # --- helper: highlight only the exact substring inside a paragraph ---
    def _highlight_exact(paragraph, token, color):
        """Highlight only the exact substring `token` inside a paragraph, preserving formatting."""
        if not token or token not in paragraph.text:
            return

        from copy import deepcopy

        def _clone_format(dst_run, src_run):
            # Copy full rPr if present (best fidelity)
            rPr = src_run._element.rPr
            if rPr is not None:
                dst_run._element.insert(0, deepcopy(rPr))
            else:
                # Fallback: copy common font attrs
                dst_run.font.name = src_run.font.name
                if src_run.font.size:
                    dst_run.font.size = src_run.font.size
                dst_run.font.bold = src_run.font.bold
                dst_run.font.italic = src_run.font.italic
                dst_run.font.underline = src_run.font.underline
                try:
                    # color may be None; guard it
                    if src_run.font.color and src_run.font.color.rgb:
                        dst_run.font.color.rgb = src_run.font.color.rgb
                except Exception:
                    pass

        i = 0
        while True:
            runs = paragraph.runs
            if i >= len(runs):
                break

            run = runs[i]
            txt = run.text or ""

            # If already highlighted token, skip
            if txt == token and getattr(run.font, "highlight_color", None) == color:
                i += 1
                continue

            pos = txt.find(token)
            if pos == -1:
                i += 1
                continue

            before = txt[:pos]
            after  = txt[pos + len(token):]

            # Keep current run as 'before' text
            run.text = before

            # Insert highlighted token run with cloned formatting
            tok_run = paragraph.add_run(token)
            _clone_format(tok_run, run)
            tok_run.font.highlight_color = color
            run._element.addnext(tok_run._element)

            # Insert trailing text (if any) with cloned formatting
            if after:
                after_run = paragraph.add_run(after)
                _clone_format(after_run, run)
                tok_run._element.addnext(after_run._element)
                i += 2  # jump past token to the 'after' run
            else:
                i += 2  # jump past token run

    # --- date highlight (only the exact date text) ---
    tomorrow = (datetime.now() + timedelta(days=1)).strftime("%d.%m.%Y")
    if isinstance(start_date, str) and start_date and start_date != tomorrow:
        for para in doc.paragraphs:
            if start_date in para.text:
                _highlight_exact(para, start_date, WD_COLOR_INDEX.RED)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if start_date in para.text:
                            _highlight_exact(para, start_date, WD_COLOR_INDEX.RED)

    # --- equipment: highlight only the word "יש" when equipment_list starts with it ---
    if isinstance(equipment_list, str) and equipment_list.strip().startswith("יש"):
        token = "יש"
        for para in doc.paragraphs:
            if token in para.text:
                _highlight_exact(para, token, WD_COLOR_INDEX.YELLOW)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if token in para.text:
                            _highlight_exact(para, token, WD_COLOR_INDEX.YELLOW)

    # --- error tokens: highlight only the exact token, not whole lines ---
    missing_data_errors = [
        "[MISSING DATA]", "[MISSING TIME]", "[MISSING LOCATION/COORDINATES]",
        "[MISSING VEHICLE DATA]", "[TARGET NOT FOUND]", "[INVALID DATE]", "[INVALID COORDS]"
    ]

    for para in doc.paragraphs:
        txt = para.text
        if not txt:
            continue
        for err in missing_data_errors:
            if err in txt:
                _highlight_exact(para, err, WD_COLOR_INDEX.RED)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    txt = para.text
                    if not txt:
                        continue
                    for err in missing_data_errors:
                        if err in txt:
                            _highlight_exact(para, err, WD_COLOR_INDEX.RED)

    doc.save("highlighted_document.docx")
    return "highlighted_document.docx"

# Streamlit UI Setup
st.title("Coordinations App")
uploaded_excel = st.file_uploader("Upload your Excel file (.xlsx, .xlsm)", type=["xlsx", "xlsm"])
user_name = st.text_input("Enter your name:")

if uploaded_excel and user_name:
    try:
        df = pd.read_excel(uploaded_excel, engine="openpyxl", header=None)
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        st.stop()

    st.dataframe(df)

    try:
        doc = DocxTemplate(TEMPLATE_FILE)
    except Exception as e:
        st.error(f"Error loading Word template: {e}")
        st.stop()

    organization_name = extract_below_target(df, "ORGANIZATION")
    raw_start_date = extract_below_target(df, "Date of coordination")
    coordination_header = extract_below_target(df, "Objective of the mission")
    start_time, end_time = extract_hours(df)
    start_date = format_and_validate_date(raw_start_date)
    formatted_levels = "\n".join(count_coordinates_levels(df))
    # Build the sites map image
    sites = extract_sites_for_map(df)
    sites_map_img = ""
    if sites:
        img_path = generate_sites_map_image(sites, out_path="sites_map.png")
        if img_path:
            sites_map_img = InlineImage(doc, img_path, width=Cm(12))  # adjust width as you like
    mission = extract_below_target(df, "Detailed execution")
    equipment_list = extract_equipment_status(df)
    table_data, white_wheels_text = extract_personnel_table_data(df)
    vehicle_table_data = extract_vehicle_table_data(df)

    data_dict = {
        "organization_name": organization_name,
        "start_date": start_date,
        "end_date": start_date,
        "start_time": start_time,
        "end_time": end_time,
        "levels": formatted_levels,
        "white_wheels": white_wheels_text,
        "soldier_name": user_name,
        "coordination_header": coordination_header,
        "mission": mission,
        "equipment_list": equipment_list,
        "sites_map": sites_map_img,
    }
    doc.render(data_dict)

    doc.save("processed_document.docx")

    final_doc = update_personnel_table("processed_document.docx", table_data, white_wheels_text)
    final_doc = update_vehicle_table(final_doc, vehicle_table_data)
    final_doc = update_highlighting(final_doc, start_date, equipment_list)

    with open(final_doc, "rb") as f:
        output_stream = BytesIO(f.read())

    st.download_button(
        label="Download Processed Word Document",
        data=output_stream,
        file_name="processed_document.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
