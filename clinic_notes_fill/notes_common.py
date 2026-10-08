"""
Shared helpers for fill_clinical_notes.py: settings, finding patient files,
loading pages, and writing the filled copy of the workbook.
"""

import datetime as dt
import json
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

import cv2
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(HERE, "settings.json")
OUT_DIR = os.path.join(HERE, "output")

TEST_LIMIT = 10
FILE_EXTS = (".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")


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


def load_pages(path):
    """Yield page images (BGR numpy arrays)."""
    low = path.lower()
    if low.endswith(".pdf"):
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


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def write_workbook(src, dst, fills):
    """fills: {row_number: {col_letter: value}}; value is text, a number, or
    ("DATE", datetime.date) - written with the date format already used in that column.
    A cell is "blank" when it is empty or holds only spaces."""
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
        shared = z.read("xl/sharedStrings.xml").decode("utf-8") if "xl/sharedStrings.xml" in names else ""
    # shared strings that are only whitespace (e.g. a non-breaking space) count as blank
    sis = re.findall(r"<si>(.*?)</si>", shared, re.S)
    blank_strings = {str(i) for i, si in enumerate(sis)
                     if "".join(re.findall(r"<t[^>]*>([^<]*)</t>", si)).replace("&#160;", " ").strip() == ""}

    def date_style(col):
        m = re.search(r'<c r="%s\d+"[^>]*?\ss="(\d+)"(?:(?!t=)[^>])*>(?:<f>[^<]*</f>)?<v>\d{5}(?:\.\d+)?</v>' % col,
                      sheet)
        return int(m.group(1)) if m else None

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
                shared_blank = 't="s"' in attrs and re.fullmatch(r"<v>(\d+)</v>", inner.strip()) \
                    and re.fullmatch(r"<v>(\d+)</v>", inner.strip()).group(1) in blank_strings
                inline_blank = re.fullmatch(r"<is>\s*<t[^>]*>\s*</t>\s*</is>", inner.strip())
                if re.search(r"<v>|<is>|<f>", inner) and not (shared_blank or inline_blank):
                    continue  # cell has a value after all - never overwrite
                sm = re.search(r'\ss="(\d+)"', attrs)
                s = int(sm.group(1)) if sm else 0
            else:
                s = 0
            if isinstance(val, tuple) and val[0] == "DATE":
                ds = date_style(col)
                if ds is None:
                    val = val[1].strftime("%m/%d/%Y")
                else:
                    val, s = (val[1] - dt.date(1899, 12, 30)).days, ds
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                new_c = f'<c r="{ref}" s="{highlighted_style(s)}"><v>{val}</v></c>'
            else:
                new_c = (f'<c r="{ref}" s="{highlighted_style(s)}" t="inlineStr">'
                         f'<is><t>{escape(str(val))}</t></is></c>')
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
