"""
Fill blank MASTER!AA:AF cells (Known CAD, Prior MI, Coronary Stents, Kidney
Disease, Smoker, Arrhythmia/AFib) from the "Patient History" checkbox table on
each patient's CCTA intake form.

Usage (normally run through the .bat files):
    python fill_ccta_intake.py            -> test run, first 10 patients
    python fill_ccta_intake.py --all      -> every patient with blanks
    python fill_ccta_intake.py --reindex  -> rescan the patient folders

Safety rules:
  * the original Excel file is never modified; a new *_FILLED.xlsx is written
  * only cells that are blank are filled; existing values are never touched
  * a cell is filled only when exactly one box (Yes or No) is clearly checked;
    anything unclear stays blank and is listed in the review log
"""

import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import sys
import time
import warnings
import zipfile
from xml.sax.saxutils import escape

import cv2
import numpy as np

warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(HERE, "settings.json")
INDEX_FILE = os.path.join(HERE, "intake_file_index.json")
OCR_CACHE_FILE = os.path.join(HERE, "ocr_cache.json")
OUT_DIR = os.path.join(HERE, "output")
CHECK_DIR = os.path.join(OUT_DIR, "check_images")

TEST_LIMIT = 10
READER_VERSION = 2  # bump when the checkbox reader changes, so old saved readings are redone
FILE_EXTS = (".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")

# Spreadsheet column -> (header text, regex that finds the row label on the form)
FIELDS = [
    ("AA", "Q: Known CAD", r"\bcad\b|artery\s*disease"),
    ("AB", "Q: Prior MI", r"history\s*of\s*mi|heart\s*attack"),
    ("AC", "Q: Coronary Stents", r"stent"),
    ("AD", "Q: Kidney Disease", r"kidney"),
    ("AE", "Q: Smoker", r"smoker"),
    ("AF", "Q: Arrhythmia/AFib", r"arrhythm|atrial\s*fib"),
]
# Other rows of the same table; read only to calibrate what an empty box looks like.
OTHER_ROWS = [r"heart\s*disease", r"valve", r"lung", r"copd|asthma"]


# --------------------------------------------------------------------------
# Settings (remembered file/folder choices)
# --------------------------------------------------------------------------

def load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1)
    os.replace(tmp, path)


def pick_paths(settings):
    """Ask (with a pop-up window) for the Excel file and the patient folder."""
    excel = settings.get("excel_file")
    root = settings.get("patient_root")
    if excel and os.path.isfile(excel) and root and os.path.isdir(root):
        return excel, root

    import tkinter as tk
    from tkinter import filedialog, messagebox
    tkroot = tk.Tk()
    tkroot.withdraw()
    tkroot.attributes("-topmost", True)
    if not (excel and os.path.isfile(excel)):
        messagebox.showinfo("Step 1 of 2", "Select the Tampa Insight Excel file.")
        excel = filedialog.askopenfilename(
            title="Select the Tampa Insight Excel file",
            filetypes=[("Excel files", "*.xlsx")])
    if not (root and os.path.isdir(root)):
        messagebox.showinfo(
            "Step 2 of 2",
            "Select the folder that contains the patient folders\n"
            "(for example S:\\ or the patients folder inside it).")
        root = filedialog.askdirectory(title="Select the patient folders location")
    tkroot.destroy()
    if not excel or not root:
        sys.exit("Cancelled - nothing selected.")
    settings["excel_file"] = excel
    settings["patient_root"] = root
    save_json(SETTINGS_FILE, settings)
    return excel, root


# --------------------------------------------------------------------------
# Read the spreadsheet: which rows need filling
# --------------------------------------------------------------------------

def read_targets(excel):
    import openpyxl
    wb = openpyxl.load_workbook(excel, read_only=True, data_only=True)
    ws = wb["MASTER"]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1, max_col=32))]
    for col, name, _ in FIELDS:
        idx = openpyxl.utils.column_index_from_string(col) - 1
        if (header[idx] or "").strip() != name:
            sys.exit(f"Column {col} header is '{header[idx]}', expected '{name}'. "
                     "The spreadsheet layout changed - stopping without changes.")
    targets = []
    for r, row in enumerate(ws.iter_rows(min_row=2, max_col=32, values_only=True), start=2):
        name = row[0]
        if not name or not str(name).strip():
            continue
        blanks = [col for col, _, _ in FIELDS
                  if row[openpyxl.utils.column_index_from_string(col) - 1] in (None, "")
                  or str(row[openpyxl.utils.column_index_from_string(col) - 1]).strip() == ""]
        if blanks:
            targets.append({"row": r, "name": str(name).strip(), "exam_date": row[1],
                            "vidistar": row[2], "dob": row[8], "blanks": blanks})
    wb.close()
    return targets


# --------------------------------------------------------------------------
# Find intake files
# --------------------------------------------------------------------------

def build_index(root):
    print(f"Scanning {root} for CCTA intake files (only needed once)...")
    files, ndirs, t0 = [], 0, time.time()
    for dirpath, dirnames, filenames in os.walk(root):
        ndirs += 1
        if ndirs % 2000 == 0:
            print(f"  ...{ndirs} folders scanned, {len(files)} intake files found")
        for fn in filenames:
            low = fn.lower()
            if low.endswith(FILE_EXTS) and "intake" in low and "ccta" in low.replace(" ", ""):
                files.append(os.path.join(dirpath, fn))
    print(f"  done: {len(files)} intake files in {ndirs} folders ({time.time() - t0:.0f}s)")
    data = {"root": root, "built": dt.datetime.now().isoformat(timespec="seconds"),
            "files": files}
    save_json(INDEX_FILE, data)
    return files


def norm_words(text):
    return re.sub(r"[^A-Z0-9]+", " ", str(text).upper()).split()


def date_variants(d):
    if not isinstance(d, (dt.date, dt.datetime)):
        return set()
    return {d.strftime(f) for f in ("%Y%m%d", "%m%d%Y", "%m%d%y", "%Y%m%d")} | {
        f"{d.month}{d.day:02d}{d.year}", f"{d.month}{d.day}{d.year % 100:02d}",
        f"{d.month}{d.day}{d.year}", f"{d.month:02d}{d.day:02d}{d.year % 100:02d}"}


def find_files(t, indexed):
    """indexed: list of (path, set_of_words, digits_only_string).

    Patient folders are named  patientID_LASTNAME_FIRSTNAME_DOB, so we look for
    (best first): ID + last name, last+first name + DOB, ID alone, name alone.
    """
    vid = str(t["vidistar"]).split(".")[0].strip() if t["vidistar"] not in (None, "") else ""
    last, _, first = t["name"].partition(",")
    last_words = [w for w in norm_words(last) if len(w) >= 2]
    first_words = [w for w in norm_words(first) if len(w) >= 2][:1]
    dob = date_variants(t.get("dob"))

    def has_id(words):
        return bool(vid) and vid in words

    def has_last(words):
        return bool(last_words) and all(w in words for w in last_words)

    def has_name(words):
        return has_last(words) and all(w in words for w in first_words)

    def has_dob(digits):
        return any(v in digits for v in dob)

    tiers = [
        ("patient ID + last name", lambda w, d: has_id(w) and has_last(w)),
        ("name + DOB", lambda w, d: has_name(w) and has_dob(d)),
        ("patient ID only (name differs - please check)", lambda w, d: has_id(w)),
        ("name only (no ID/DOB match - please check)", lambda w, d: has_name(w)),
    ]
    for how, test in tiers:
        cands = [p for p, words, digits in indexed if test(words, digits)]
        if cands:
            if len(cands) > 1:
                dv = date_variants(t["exam_date"])
                dated = [p for p, _, digits in indexed if p in cands and any(v in digits for v in dv)]
                if dated:
                    cands, how = dated, how + " + exam date"
            return cands, how
    return [], ""


# --------------------------------------------------------------------------
# OCR + checkbox reading
# --------------------------------------------------------------------------

_ocr = None


def ocr(img):
    global _ocr
    if _ocr is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr = RapidOCR()
    result, _ = _ocr(img)
    tokens = []
    for box, text, _score in result or []:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        tokens.append({"x0": min(xs), "x1": max(xs), "y0": min(ys), "y1": max(ys),
                       "t": str(text)})
    return tokens


def load_pages(path):
    """Yield page images (BGR numpy arrays)."""
    low = path.lower()
    if low.endswith(".pdf"):
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        with fitz.open(path) as doc:
            for page in doc:
                pix = page.get_pixmap(dpi=200)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
                yield cv2.cvtColor(img, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR) \
                    if pix.n in (3, 4) else cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    else:
        data = np.fromfile(path, dtype=np.uint8)  # works with any characters in the path
        if low.endswith((".tif", ".tiff")):
            ok, pages = cv2.imdecodemulti(data, cv2.IMREAD_COLOR)
            for p in (pages if ok else []):
                yield p
        else:
            img = cv2.imdecode(data, cv2.IMREAD_COLOR)
            if img is not None:
                yield img


def find_table(tokens):
    """Return header token of the Patient History table, or None."""
    for tk in tokens:
        if "patienthistory" in re.sub(r"[^a-z]", "", tk["t"].lower()):
            return tk
    return None


def read_form(path):
    """Read the Patient History table. Returns dict with values per column + diagnostics."""
    for page_no, img in enumerate(load_pages(path), start=1):
        # Large scans: shrink so OCR stays fast; keep enough resolution for the boxes.
        scale = 1.0
        if max(img.shape[:2]) > 2600:
            scale = 2600 / max(img.shape[:2])
            img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        tokens = ocr(img)
        header = find_table(tokens)
        if header is None:
            continue
        res = read_table(img, tokens, header)
        res["page"] = page_no
        return res
    return {"error": "Patient History table not found"}


def read_table(img, tokens, header):
    h_line = header["y1"] - header["y0"]
    # The form has other tables (Surgical History, Other History...) to the
    # right of Patient History. Find where the next table starts so that we
    # only ever look inside the Patient History table.
    right_limit = float("inf")
    for t in tokens:
        if t is header or t["x0"] <= header["x1"]:
            continue
        near_header_line = abs((t["y0"] + t["y1"]) / 2 - (header["y0"] + header["y1"]) / 2) < 2.5 * h_line
        if near_header_line and re.search(r"histor|surg", t["t"].lower()):
            right_limit = min(right_limit, t["x0"] - 0.5 * h_line)
    # Tokens that belong to the table: below the header, within ~16 lines.
    below = [t for t in tokens
             if t["y0"] >= header["y0"] + 0.5 * h_line and t["y0"] <= header["y1"] + 16 * h_line * 1.9
             and t["x1"] >= header["x0"] - 2 * h_line and t["x0"] < right_limit]

    def label_row(pattern):
        for t in sorted(below, key=lambda t: t["y0"]):
            txt = t["t"].lower()
            # row labels are left-aligned under the "Patient History" header
            if t["x0"] > header["x0"] + 3 * h_line:
                continue
            if re.search(pattern, txt) and not re.search(r"\byes\b", txt):
                # avoid matching the "Coronary Stents" row for CAD's "coronary"
                return t
        return None

    rows = {}
    for col, _, pat in FIELDS:
        rows[col] = label_row(pat)
    others = [label_row(p) for p in OTHER_ROWS]

    labels = [r for r in list(rows.values()) + others if r is not None]
    if not labels:
        return {"error": "table labels not found"}
    label_right = max(t["x1"] for t in labels)
    row_h = float(np.median([t["y1"] - t["y0"] for t in labels]))

    # Yes / No tokens to the right of the labels define the two checkbox columns.
    yes_tok = [t for t in below if t["x0"] > label_right - row_h and re.search(r"y\W*e?\W*s", t["t"].lower())]
    no_tok = [t for t in below if t["x0"] > label_right - row_h
              and re.search(r"n\W*[o0]\W*$", t["t"].lower()) and not re.search(r"yes", t["t"].lower())]
    if not yes_tok or not no_tok:
        return {"error": f"Yes/No columns not found (yes={len(yes_tok)}, no={len(no_tok)})"}
    yes_x = float(np.median([t["x0"] for t in yes_tok]))
    no_x = float(np.median([t["x0"] for t in no_tok]))
    if no_x <= yes_x + row_h:
        return {"error": "Yes/No columns overlap"}

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # Ink threshold from the table area (handles grey/dark scans).
    ty0 = int(max(0, header["y0"]))
    ty1 = int(min(gray.shape[0], max(t["y1"] for t in labels) + row_h))
    tx0 = int(max(0, header["x0"] - row_h))
    tx1 = int(min(gray.shape[1], no_x + 4 * row_h))
    thr, _ = cv2.threshold(gray[ty0:ty1, tx0:tx1], 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    thr = min(thr, 200)

    # Pin down the boxes using the table's printed vertical grid lines:
    #   | label | [] Yes | [] No |
    # The OCR word positions are only a first guess (a check mark drawn over
    # a box can make the OCR skip the box).
    box_pos = "from OCR words"
    gx0 = int(max(0, label_right - 2 * row_h))
    gx1 = int(min(gray.shape[1], right_limit if right_limit != float("inf") else no_x + 6 * row_h))
    if gx1 - gx0 > 4 * row_h and ty1 - ty0 > 4 * row_h:
        dark = (gray[ty0:ty1, gx0:gx1] < thr).astype(np.uint8)
        dark_w = cv2.dilate(dark, np.ones((1, 3), np.uint8))  # tolerate slight tilt
        vk = max(10, int(2 * row_h))
        vert = cv2.morphologyEx(dark_w, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, vk)))
        cols = np.where(vert.sum(axis=0) > 0.4 * (ty1 - ty0))[0]
        lines = []
        for x in cols:
            if lines and x - lines[-1][-1] <= 2:
                lines[-1].append(x)
            else:
                lines.append([x])
        lines = [gx0 + float(np.mean(l)) for l in lines]
        l1 = [x for x in lines if x <= yes_x + 0.3 * row_h]
        l2 = [x for x in lines if x > yes_x + row_h]
        if l1 and l2:
            L1, L2 = max(l1), min(l2)
            # left edge of the printed Yes box = first strong dark column in the Yes cell
            cell = dark[:, int(L1 - gx0 + 3):int(L2 - gx0 - 3)]
            if cell.size:
                prof = cell.sum(axis=0)
                strong = np.where(prof >= 0.5 * prof.max())[0]
                if len(strong):
                    box_left = L1 + 3 + float(strong[0])
                    if box_left < L1 + 0.6 * (L2 - L1):
                        yes_x = box_left
                        no_x = L2 + (box_left - L1)
                        box_pos = "from table grid lines"

    # distance between table rows, so a box area never reaches into the row above
    ycs = sorted((t["y0"] + t["y1"]) / 2 for t in labels)
    gaps = [b - a for a, b in zip(ycs, ycs[1:]) if b - a > 0.8 * row_h]
    pitch = float(np.median(gaps)) if gaps else 2 * row_h

    def box_rect(x_left, lab):
        # the box itself plus room above/right for check marks that sweep out of it
        yc = (lab["y0"] + lab["y1"]) / 2
        up = min(1.4 * row_h, 0.5 * pitch)
        down = min(0.65 * row_h, 0.45 * pitch)
        return (int(x_left - 0.2 * row_h), int(yc - up),
                int(x_left + 1.15 * row_h), int(yc + down))

    f = max(1.0, 40.0 / row_h)  # work at ~40 px per text line

    def dark_crop(rect):
        x0, y0, x1, y1 = rect
        crop = gray[max(0, y0):max(0, y1), max(0, x0):max(0, x1)]
        if crop.size == 0:
            return None
        crop = cv2.resize(crop, None, fx=f, fy=f, interpolation=cv2.INTER_CUBIC)
        # darkness 0..1, with paper/background noise (lighter than the ink threshold) removed
        d = (255.0 - crop.astype(np.float32)) / 255.0
        paper = (255.0 - min(255.0, thr + 0.5 * (255 - thr))) / 255.0
        d = np.clip((d - paper) / max(1e-3, 1 - paper), 0, 1)
        # erase the table's long printed grid lines (longer than any box edge)
        b = (d > 0.4).astype(np.uint8)
        k = max(8, int(1.5 * row_h * f))
        grid = cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (k, 1)))
        grid |= cv2.morphologyEx(b, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, k)))
        grid = cv2.dilate(grid, np.ones((5, 5), np.uint8))
        d[grid > 0] = 0
        return d

    all_rows = [r for r in list(rows.values()) + others if r is not None]
    crops = {}
    for lab in all_rows:
        for x in (yes_x, no_x):
            c = dark_crop(box_rect(x, lab))
            if c is not None:
                crops[(id(lab), x)] = c
    if not crops:
        return {"error": "could not cut out the boxes"}
    # The cleanest box on the form = what an EMPTY printed box looks like here.
    template = min(crops.values(), key=lambda c: c.mean())
    template_fat = cv2.dilate(template, np.ones((3, 3), np.uint8))
    shift = max(2, int(0.2 * row_h * f))

    def ink(lab, x):
        """Share of the box area covered by pen marks that are not part of the
        printed empty box (best alignment against the empty-box template)."""
        c = crops.get((id(lab), x))
        if c is None:
            return 0.0
        h = min(c.shape[0], template.shape[0])
        w = min(c.shape[1], template.shape[1])
        c = c[:h, :w]
        best = None
        for dy in range(-shift, shift + 1, 2):
            for dx in range(-shift, shift + 1, 2):
                t = np.roll(np.roll(template_fat[:h, :w], dy, axis=0), dx, axis=1)
                extra = float(np.clip(c - t, 0, 1).mean())
                best = extra if best is None else min(best, extra)
        return best

    inks = [ink(lab, x) for lab in all_rows for x in (yes_x, no_x)]
    # At least half the boxes on a form are empty, so the lower half tells us
    # how much "noise" an empty box shows on this particular scan.
    low = sorted(inks)[: max(2, len(inks) // 2)]
    empty = float(np.median(low))
    noise = 1.4826 * float(np.median(np.abs(np.array(low) - empty)))  # robust to a stray mark
    gap = max(0.008, 4 * noise)

    out = {"values": {}, "detail": {}, "rects": [], "crop": (tx0, ty0, tx1, ty1),
           "empty_ink": round(empty, 4), "box_pos": box_pos}
    for col, _, _ in FIELDS:
        lab = rows[col]
        if lab is None:
            out["values"][col] = None
            out["detail"][col] = "row not found"
            continue
        ry, rn = box_rect(yes_x, lab), box_rect(no_x, lab)
        iy, i_n = ink(lab, yes_x), ink(lab, no_x)
        # A box counts as checked when it clearly has more pen ink than an empty
        # box AND clearly more than the other box in the same row.
        if iy >= empty + gap and iy >= i_n + gap and iy >= 2 * i_n:
            val = "Yes"
        elif i_n >= empty + gap and i_n >= iy + gap and i_n >= 2 * iy:
            val = "No"
        else:
            val = None
        out["values"][col] = val
        out["detail"][col] = f"yes={iy:.3f} no={i_n:.3f}" + ("" if val else " UNCLEAR")
        out["rects"].append((col, ry, rn, val))
    return out


def save_check_image(path_out, img_path, res):
    """Save a picture of the table with the boxes the script looked at."""
    if "crop" not in res:
        return
    for page_no, img in enumerate(load_pages(img_path), start=1):
        if page_no != res.get("page"):
            continue
        if max(img.shape[:2]) > 2600:
            s = 2600 / max(img.shape[:2])
            img = cv2.resize(img, None, fx=s, fy=s, interpolation=cv2.INTER_AREA)
        for col, ry, rn, val in res["rects"]:
            colr = {"Yes": (0, 160, 0), "No": (200, 0, 0), None: (0, 0, 255)}[val]
            cv2.rectangle(img, ry[:2], ry[2:], colr, 2)
            cv2.rectangle(img, rn[:2], rn[2:], colr, 2)
            cv2.putText(img, f"{col}:{val or '??'}", (rn[2] + 5, rn[3]),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, colr, 1)
        x0, y0, x1, y1 = res["crop"]
        crop = img[max(0, y0 - 10):y1 + 10, max(0, x0 - 10):x1 + 120]
        ok, buf = cv2.imencode(".png", crop)
        if ok:
            buf.tofile(path_out)
        return


# --------------------------------------------------------------------------
# Write the new Excel file (edits the file's XML directly so that dropdowns,
# formatting and everything else in the workbook are kept exactly as they are)
# --------------------------------------------------------------------------

def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def write_workbook(src, dst, fills):
    """fills: {row_number: {col_letter: 'Yes'/'No'}}"""
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        wb_xml = z.read("xl/workbook.xml").decode("utf-8")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        rid = re.search(r'<sheet [^>]*name="MASTER"[^>]*r:id="([^"]+)"', wb_xml).group(1)
        target = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) \
            or re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
        sheet_path = "xl/" + target.group(1).lstrip("/").replace("xl/", "", 1)
        sheet = z.read(sheet_path).decode("utf-8")
        styles = z.read("xl/styles.xml").decode("utf-8")

    # --- styles: add a light-yellow fill and highlighted copies of used cell styles
    fills_m = re.search(r'<fills count="(\d+)">(.*?)</fills>', styles, re.S)
    new_fill_id = int(fills_m.group(1))
    styles = styles.replace(
        fills_m.group(0),
        f'<fills count="{new_fill_id + 1}">{fills_m.group(2)}'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/>'
        '<bgColor indexed="64"/></patternFill></fill></fills>')
    xfs_m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    xf_list = re.findall(r'<xf\b[^>]*?/>|<xf\b[^>]*?>.*?</xf>', xfs_m.group(2), re.S)
    highlight = {}  # original style index -> new style index

    def highlighted_style(s):
        if s not in highlight:
            xf = xf_list[s]
            xf = re.sub(r'\sfillId="\d+"', "", xf, count=1)
            xf = re.sub(r'\sapplyFill="\d"', "", xf, count=1)
            xf = xf.replace("<xf", f'<xf fillId="{new_fill_id}" applyFill="1"', 1)
            highlight[s] = len(xf_list) + len(highlight)
            new_xfs.append(xf)
        return highlight[s]

    new_xfs = []

    # --- sheet: put values into the blank cells
    def fix_row(m):
        rnum = int(m.group(1))
        if rnum not in fills:
            return m.group(0)
        row = m.group(0)
        if row.endswith("/>") and "</row>" not in row:  # empty self-closing row
            row = row[:-2] + "></row>"
        for col, val in fills[rnum].items():
            ref = f"{col}{rnum}"
            cm = re.search(r'<c r="%s"(?=[\s/>])([^>]*?)(/>|>(.*?)</c>)' % ref, row, re.S)
            if cm:
                attrs = cm.group(1)
                inner = cm.group(3) or ""
                if re.search(r"<v>|<is>|<f>", inner):
                    continue  # cell has a value after all - never overwrite
                sm = re.search(r'\ss="(\d+)"', attrs)
                s = int(sm.group(1)) if sm else 0
            else:
                s = 0
            new_c = (f'<c r="{ref}" s="{highlighted_style(s)}" t="inlineStr">'
                     f'<is><t>{escape(val)}</t></is></c>')
            if cm:
                row = row[:cm.start()] + new_c + row[cm.end():]
            else:
                # insert in column order
                pos = None
                for cc in re.finditer(r'<c r="([A-Z]+)\d+"', row):
                    if col_num(cc.group(1)) > col_num(col):
                        pos = cc.start()
                        break
                if pos is None:
                    pos = row.rindex("</row>")
                row = row[:pos] + new_c + row[pos:]
        return row

    sheet = re.sub(r'<row r="(\d+)"[^>]*?(?:/>|>.*?</row>)', fix_row, sheet, flags=re.S)
    styles = styles.replace(
        xfs_m.group(0),
        f'<cellXfs count="{len(xf_list) + len(new_xfs)}">{xfs_m.group(2)}{"".join(new_xfs)}</cellXfs>')

    tmp = dst + ".tmp"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == sheet_path:
                zout.writestr(item, sheet.encode("utf-8"))
            elif item.filename == "xl/styles.xml":
                zout.writestr(item, styles.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    os.replace(tmp, dst)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--all", action="store_true", help="process every patient (default: first 10)")
    ap.add_argument("--reindex", action="store_true", help="rescan the patient folders")
    ap.add_argument("--choose", action="store_true", help="choose the Excel file / folder again")
    args = ap.parse_args()

    settings = load_json(SETTINGS_FILE, {})
    if args.choose:
        settings = {}
    excel, root = pick_paths(settings)
    os.makedirs(CHECK_DIR, exist_ok=True)

    print(f"Excel file : {excel}")
    print(f"Patients   : {root}")
    targets = read_targets(excel)
    print(f"{len(targets)} patients have blank cells in AA-AF.")
    if not args.all:
        targets = targets[:TEST_LIMIT]
        print(f"TEST RUN: only the first {len(targets)} patients.")

    index = load_json(INDEX_FILE, {})
    if args.reindex or index.get("root") != root:
        files = build_index(root)
    else:
        files = index["files"]
        print(f"Using saved list of {len(files)} intake files (run with --reindex to rescan).")
    indexed = [(p, set(norm_words(os.path.relpath(p, root))),
                re.sub(r"\D", "", os.path.relpath(p, root))) for p in files]

    cache = load_json(OCR_CACHE_FILE, {})
    fills, log_rows = {}, []
    stats = {"filled_cells": 0, "patients_no_file": 0, "unclear_cells": 0}
    for i, t in enumerate(targets, start=1):
        print(f"[{i}/{len(targets)}] row {t['row']}: {t['name']}", end=" ... ", flush=True)
        cands, how = find_files(t, indexed)
        log = {"Excel row": t["row"], "Name": t["name"], "VIDISTAR ID": t["vidistar"],
               "Exam date": t["exam_date"].strftime("%m/%d/%Y")
               if isinstance(t["exam_date"], (dt.date, dt.datetime)) else t["exam_date"],
               "Matched by": how, "Intake file": "", "Status": ""}
        if not cands:
            stats["patients_no_file"] += 1
            log["Status"] = "NO INTAKE FILE FOUND"
            print("no intake file found")
            log_rows.append(log)
            continue
        if len(cands) > 4:
            log["Status"] = f"TOO MANY MATCHING FILES ({len(cands)}) - check manually"
            print(log["Status"])
            log_rows.append(log)
            continue

        results = []
        for path in cands:
            key = f"v{READER_VERSION}|{path}|{os.path.getmtime(path) if os.path.exists(path) else 0}"
            if key not in cache:
                try:
                    res = read_form(path)
                except Exception as e:  # damaged file etc.
                    res = {"error": f"could not read file: {e}"}
                cache[key] = {k: v for k, v in res.items() if k in ("values", "detail", "error", "page", "empty_ink")}
                save_json(OCR_CACHE_FILE, cache)
                if "values" in res:
                    safe = re.sub(r"[^A-Za-z0-9]+", "_", t["name"])[:40]
                    save_check_image(os.path.join(CHECK_DIR, f"row{t['row']}_{safe}_{len(results)}.png"),
                                     path, res)
            results.append((path, cache[key]))

        good = [(p, r) for p, r in results if "values" in r]
        log["Intake file"] = " | ".join(p for p, _ in results)
        if not good:
            log["Status"] = "; ".join(r.get("error", "") for _, r in results)
            print(log["Status"])
            log_rows.append(log)
            continue

        row_fill, notes = {}, []
        for col, name, _ in FIELDS:
            vals = {r["values"].get(col) for _, r in good}
            details = " / ".join(r["detail"].get(col, "") for _, r in good)
            v = vals.pop() if len(vals) == 1 else None
            if len(good) > 1 and len({r["values"].get(col) for _, r in good}) > 1:
                details += " FILES DISAGREE"
            log[name] = v or ""
            log[name + " (scores)"] = details
            if col in t["blanks"]:
                if v:
                    row_fill[col] = v
                    stats["filled_cells"] += 1
                else:
                    stats["unclear_cells"] += 1
                    notes.append(f"{col} unclear")
            else:
                log[name] += " (already had a value - not changed)"
        if row_fill:
            fills[t["row"]] = row_fill
        log["Status"] = "OK" if not notes else "PARTIAL: " + ", ".join(notes)
        print(log["Status"], {c: v for c, v in row_fill.items()})
        log_rows.append(log)

    # --- outputs
    base = os.path.splitext(os.path.basename(excel))[0]
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
    tag = "ALL" if args.all else "TEST"
    out_xlsx = os.path.join(OUT_DIR, f"{base}_FILLED_{tag}_{stamp}.xlsx")
    out_log = os.path.join(OUT_DIR, f"review_log_{tag}_{stamp}.csv")
    if fills:
        write_workbook(excel, out_xlsx, fills)
    cols = ["Excel row", "Name", "VIDISTAR ID", "Exam date", "Status", "Matched by", "Intake file"]
    for _, name, _ in FIELDS:
        cols += [name, name + " (scores)"]
    with open(out_log, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(log_rows)

    print("\n================ SUMMARY ================")
    print(f"Patients processed       : {len(targets)}")
    print(f"Cells filled             : {stats['filled_cells']}")
    print(f"Cells left blank (unclear): {stats['unclear_cells']}")
    print(f"Patients with no file    : {stats['patients_no_file']}")
    print(f"Filled Excel (yellow = new): {out_xlsx if fills else '(none - nothing to fill)'}")
    print(f"Review log               : {out_log}")
    print(f"Pictures to spot-check   : {CHECK_DIR}")
    print("Your original Excel file was NOT changed.")


if __name__ == "__main__":
    main()
