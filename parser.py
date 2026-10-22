"""Streamlit app: Indonesian legal document (PDF/DOCX) -> structured review table.

Single-file version. Deploy with:  app.py + requirements.txt
Sections: 1) patterns & hierarchy detector  2) PDF reader  3) DOCX reader
          4) table builder & Excel export   5) Streamlit UI
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from io import BytesIO

import pandas as pd
import pdfplumber
import regex as re
import streamlit as st
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter



# ======================================================================
# 1. HIERARCHY DETECTOR
# ======================================================================
# ----------------------------------------------------------------- patterns
RE_BAB = re.compile(r"^BAB\s+([IVXLCDM]+)\s*[:.]?\s*(.*)$")
RE_BAGIAN = re.compile(r"^(Bagian|Paragraf)\s+(Ke\p{L}+|Pertama)\s*$")
RE_PASAL = re.compile(r"^Pasal\s+(\d+\s?[A-Z]?)\s*$")
RE_AYAT = re.compile(r"^\(\s*(\d{1,3}[a-z]?)\s*\)\s*(.*)$")
RE_HURUF = re.compile(r"^([a-z])[.)](?:\s+(.*))?$")
RE_ANGKA = re.compile(r"^(\d{1,3})\.(?:\s+(.*))?$")
RE_MENIMBANG = re.compile(r"^Menimbang\s*:?\s*(.*)$")
RE_MENGINGAT = re.compile(r"^Mengingat\s*:?\s*(.*)$")
RE_MENETAPKAN = re.compile(r"^Menetapkan\s*:?\s*(.*)$")
RE_MEMUTUSKAN = re.compile(r"^MEMUTUSKAN\s*:?\s*$")
RE_STOP = re.compile(r"^(PENJELASAN|LAMPIRAN)\b")
RE_PAGE_NUM = re.compile(
    r"^\s*(?:[-–—]\s*)?(?:hal(?:aman)?\.?\s*|page\s*)?\d{1,4}"
    r"\s*(?:[-–—]|/\s*\d+|of\s+\d+|dari\s+\d+)?\s*$", re.I)


def is_page_number(line: str) -> bool:
    """True for lines such as '3', '- 3 -', 'Halaman 3', '3/25', 'Page 3 of 25'."""
    return bool(RE_PAGE_NUM.match(line))


# --------------------------------------------------------------- data model
@dataclass
class Item:
    section: str                      # Menimbang | Mengingat | Menetapkan | Body
    bab: str = ""
    pasal: str = ""
    path: list = field(default_factory=list)   # [("ayat","(2)"),("huruf","a")]
    lines: list = field(default_factory=list)
    has_children: bool = False

    @property
    def text(self) -> str:
        return re.sub(r"\s+", " ", " ".join(self.lines)).strip()


@dataclass
class ParsedDocument:
    items: list = field(default_factory=list)
    bab_titles: dict = field(default_factory=dict)


# ------------------------------------------------------------------- parser
class HierarchyParser:
    def __init__(self):
        self.doc = ParsedDocument()
        self.mode = None          # None|menimbang|mengingat|menetapkan|body
        self.bab = self.pasal = ""
        self.stack: list = []
        self.cur: Item | None = None
        self.expect_title = None  # "bab" | "bagian" -> next line may be a title
        self.done = False

    # -- helpers
    def _new(self, section, path):
        it = Item(section, self.bab, self.pasal, list(path))
        self.doc.items.append(it)
        self.cur = it
        return it

    def _is_marker(self, ln):
        return any(r.match(ln) for r in (RE_AYAT, RE_HURUF, RE_ANGKA, RE_PASAL, RE_BAB))

    def _push(self, typ, val, text):
        if typ == "ayat":
            self.stack = [("ayat", f"({val})")]
        else:
            types = [t for t, _ in self.stack]
            if typ in types:                       # sibling -> replace, drop deeper levels
                self.stack = self.stack[: types.index(typ)] + [(typ, val)]
            else:                                  # first child of current level
                self.stack.append((typ, val))
        self._new("Body", self.stack)
        if text:
            self.cur.lines.append(text)

    # -- main entry per line
    def feed(self, ln: str):
        if self.done or not ln:
            return
        if self.mode in ("menetapkan", "body") and RE_STOP.match(ln):
            self.done = True
            return
        if RE_MEMUTUSKAN.match(ln):
            return
        # section keywords
        if self.mode is None and (m := RE_MENIMBANG.match(ln)):
            self.mode = "menimbang"; self.cur = None
            return self.feed(m.group(1).strip())
        if self.mode in (None, "menimbang") and (m := RE_MENGINGAT.match(ln)):
            self.mode = "mengingat"; self.cur = None
            return self.feed(m.group(1).strip())
        if self.mode in (None, "menimbang", "mengingat") and (m := RE_MENETAPKAN.match(ln)):
            self.mode = "menetapkan"; self._new("Menetapkan", [])
            return self.feed(m.group(1).strip())
        if self.mode is None:
            return                                  # preamble (title, "Dengan Rahmat...")

        # structural headings
        if self.mode in ("menetapkan", "body"):
            if m := RE_BAB.match(ln):
                self.mode = "body"; self.bab = f"BAB {m.group(1)}"
                self.pasal = ""; self.stack = []; self.cur = None
                self.doc.bab_titles[self.bab] = m.group(2).strip()
                self.expect_title = None if m.group(2) else "bab"
                return
            if RE_BAGIAN.match(ln):
                self.expect_title = "bagian"; self.cur = None
                return
            if m := RE_PASAL.match(ln):
                self.mode = "body"; self.pasal = f"Pasal {m.group(1).replace(' ', '')}"
                self.stack = []; self._new("Body", [])
                self.expect_title = None
                return
            if self.expect_title:                   # heading title line(s)
                if not self._is_marker(ln) and len(ln) < 120 and not ln.endswith((".", ":", ";")):
                    if self.expect_title == "bab":
                        t = self.doc.bab_titles.get(self.bab, "")
                        self.doc.bab_titles[self.bab] = (t + " " + ln).strip()
                    return
                self.expect_title = None

        # list markers
        if self.mode == "menimbang":
            if m := RE_HURUF.match(ln):
                self._new("Menimbang", [("huruf", m.group(1))]); self.cur.lines.append(m.group(2) or "")
                return
        elif self.mode == "mengingat":
            if m := RE_ANGKA.match(ln):
                self._new("Mengingat", [("angka", m.group(1))]); self.cur.lines.append(m.group(2) or "")
                return
        elif self.mode == "body":
            for rx, typ in ((RE_AYAT, "ayat"), (RE_HURUF, "huruf"), (RE_ANGKA, "angka")):
                if m := rx.match(ln):
                    return self._push(typ, m.group(1), (m.group(2) or "").strip())

        # continuation line -> merge into current provision
        if self.cur is None:
            self._new({"menimbang": "Menimbang", "mengingat": "Mengingat",
                       "menetapkan": "Menetapkan"}.get(self.mode, "Body"), [])
        self.cur.lines.append(ln)

    def finish(self) -> ParsedDocument:
        items = self.doc.items
        for i, it in enumerate(items[:-1]):         # mark parents (have deeper children)
            nxt = items[i + 1]
            n = len(it.path)
            if (nxt.section == it.section and nxt.pasal == it.pasal and nxt.bab == it.bab
                    and len(nxt.path) > n and nxt.path[:n] == it.path):
                it.has_children = True
        return self.doc


def detect_hierarchy(lines: list[str]) -> ParsedDocument:
    """Run the state machine over cleaned lines."""
    p = HierarchyParser()
    for ln in lines:
        ln = ln.strip()
        if is_page_number(ln):          # safety net: drop stray page numbers
            continue
        p.feed(ln)
    return p.finish()

# ======================================================================
# 2. PDF READER
# ======================================================================
def _norm(line: str) -> str:
    return re.sub(r"\d+", "#", line.lower()).strip()


def extract_pdf_lines(file) -> tuple[list[str], list[str]]:
    """Return (lines, warnings). `file` is a path or file-like object."""
    pages: list[list[str]] = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            pages.append([l.strip() for l in txt.splitlines() if l.strip()])

    warnings = []
    if not any(pages):
        return [], ["PDF tidak memiliki teks (kemungkinan hasil scan). Gunakan OCR terlebih dahulu."]

    # Repeated lines at page edges (top/bottom 2 lines) = header/footer
    counter: Counter = Counter()
    for pl in pages:
        counter.update({_norm(l) for l in pl[:2] + pl[-2:]})
    threshold = max(3, int(len(pages) * 0.4))
    repeated = {k for k, v in counter.items() if v >= threshold} if len(pages) >= 3 else set()

    lines: list[str] = []
    for pl in pages:
        last = len(pl) - 1
        for i, l in enumerate(pl):
            at_edge = i < 2 or i > last - 2
            if at_edge and (_norm(l) in repeated or is_page_number(l)):
                continue
            lines.append(l)
    return lines, warnings

# ======================================================================
# 3. DOCX READER
# ======================================================================
def _iter_paragraphs(doc):
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            for row in Table(child, doc).rows:
                for cell in row.cells:
                    yield from cell.paragraphs


def extract_docx_lines(file) -> tuple[list[str], list[str]]:
    """Headers/footers live outside the body, so they are excluded automatically."""
    doc = Document(file)
    lines, auto_numbered = [], 0
    for p in _iter_paragraphs(doc):
        txt = p.text.strip()
        if not txt or is_page_number(txt):
            continue
        has_num = p._p.pPr is not None and p._p.pPr.numPr is not None
        if has_num and not any(r.match(txt) for r in (RE_AYAT, RE_HURUF, RE_ANGKA)):
            auto_numbered += 1
        lines.extend(l.strip() for l in txt.splitlines() if l.strip())
    warnings = []
    if auto_numbered:
        warnings.append(
            f"{auto_numbered} paragraf memakai penomoran otomatis Word (nomor tidak tersimpan sebagai teks). "
            "Ubah ke teks biasa (Select all > Home > Numbering off, atau simpan ke PDF) agar ayat/huruf/angka terdeteksi.")
    return lines, warnings

# ======================================================================
# 4. TABLE BUILDER & EXCEL
# ======================================================================
COLUMNS = ["No", "Bab", "Pasal", "Ayat", "Huruf", "Angka", "Bagian",
           "Draft RPM", "Masukan", "Justifikasi"]


def _label(item) -> str:
    if item.section == "Menimbang":
        return f"Menimbang Huruf {item.path[0][1]}" if item.path else "Menimbang"
    if item.section == "Mengingat":
        return f"Mengingat Poin {item.path[0][1]}" if item.path else "Mengingat"
    if item.section == "Menetapkan":
        return "Menetapkan"
    parts = [item.pasal] + [f"Ayat {v}" if t == "ayat" else f"{t} {v}" for t, v in item.path]
    return " ".join(p for p in parts if p)


def build_rows(doc: ParsedDocument, include_parent_text: bool = False) -> pd.DataFrame:
    """One row per leaf provision. Parents that only introduce sub-items
    (e.g. '(3) meliputi:') are skipped unless include_parent_text=True."""
    rows = []
    for it in doc.items:
        if not it.text or (it.has_children and not include_parent_text):
            continue
        p = dict(it.path)
        rows.append({
            "Bab": it.bab if it.section == "Body" else "",
            "Pasal": it.pasal if it.section == "Body" else "",
            "Ayat": p.get("ayat", ""),
            "Huruf": p.get("huruf", ""),
            "Angka": p.get("angka", ""),
            "Bagian": _label(it),
            "Draft RPM": it.text,
            "Masukan": "", "Justifikasi": "",
        })
    df = pd.DataFrame(rows, columns=COLUMNS[1:])
    df.insert(0, "No", range(1, len(df) + 1))
    return df


def hierarchy_preview(doc: ParsedDocument, width: int = 90) -> str:
    """Indented text tree for the UI."""
    out, last_bab, last_pasal = [], None, None
    for it in doc.items:
        if it.section != "Body":
            out.append(f"{_label(it)}: {it.text[:width]}")
            continue
        if it.bab != last_bab:
            out.append(f"\n{it.bab} {doc.bab_titles.get(it.bab, '')}".rstrip())
            last_bab, last_pasal = it.bab, None
        if it.pasal != last_pasal:
            out.append(f"  {it.pasal}")
            last_pasal = it.pasal
        if it.path:
            tag = " ".join(f"{t}:{v}" if t != "ayat" else v for t, v in it.path)
            out.append("    " + "  " * (len(it.path) - 1) + f"{tag}  {it.text[:width]}")
        elif it.text:
            out.append(f"    {it.text[:width]}")
    return "\n".join(out)


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Review RPM")
        ws = xw.sheets["Review RPM"]
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
            c.alignment = Alignment(horizontal="center", vertical="center")
        widths = {"No": 6, "Bab": 10, "Pasal": 10, "Ayat": 8, "Huruf": 8, "Angka": 8,
                  "Bagian": 28, "Draft RPM": 70, "Masukan": 40, "Justifikasi": 40}
        for i, name in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(name, 15)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()

# ======================================================================
# 5. STREAMLIT UI
# ======================================================================
def main():
    st.set_page_config(page_title="Parser RPM", layout="wide")
    st.title("📑 Parser Draft Regulasi → Tabel Review")

    with st.sidebar:
        st.header("Opsi")
        include_parent = st.checkbox("Sertakan baris induk (kalimat pengantar, mis. 'meliputi:')", False)

    up = st.file_uploader("Upload dokumen (PDF atau DOCX)", type=["pdf", "docx"])
    if not up:
        st.info("Unggah file PDF atau DOCX untuk memulai.")
        st.stop()

    try:
        with st.spinner("Membaca dokumen..."):
            if up.name.lower().endswith(".pdf"):
                lines, warns = extract_pdf_lines(up)
            else:
                lines, warns = extract_docx_lines(up)
            doc = detect_hierarchy(lines)
            df = build_rows(doc, include_parent)
    except Exception as e:  # noqa: BLE001
        st.error(f"Gagal memproses file: {e}")
        st.stop()

    for w in warns:
        st.warning(w)
    if df.empty:
        st.error("Tidak ada struktur terdeteksi. Pastikan dokumen memuat Menimbang/Mengingat/Menetapkan/BAB/Pasal.")
        st.stop()

    c1, c2, c3 = st.columns(3)
    c1.metric("Baris", len(df))
    c2.metric("BAB", len(doc.bab_titles))
    c3.metric("Pasal", df["Pasal"].replace("", None).nunique())

    tab1, tab2, tab3 = st.tabs(["Hierarki", "Tabel (dapat diedit)", "Teks mentah"])
    with tab1:
        st.code(hierarchy_preview(doc), language="text")
    with tab2:
        edited = st.data_editor(df, use_container_width=True, hide_index=True, num_rows="dynamic",
                                disabled=["No", "Bab", "Pasal", "Ayat", "Huruf", "Angka", "Bagian", "Draft RPM"])
        st.download_button("⬇️ Download Excel", to_excel_bytes(edited),
                           file_name=up.name.rsplit(".", 1)[0] + "_review.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with tab3:
        st.text_area("Baris hasil ekstraksi (setelah pembersihan)", "\n".join(lines), height=400)


if __name__ == "__main__":
    main()
