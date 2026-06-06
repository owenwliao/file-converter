import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor DPI aware
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import io
import csv
import json
import os
import sys
import threading
import zipfile
import xml.etree.ElementTree as ET
from xml.dom import minidom
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog


# ── conversion map ────────────────────────────────────────────────────────────

CONVERSION_MAP = {
    "png":  ["jpg", "webp", "bmp", "gif", "tiff", "ico"],
    "jpg":  ["png", "webp", "bmp", "gif", "tiff", "ico"],
    "jpeg": ["png", "webp", "bmp", "gif", "tiff", "ico"],
    "webp": ["png", "jpg", "bmp", "gif", "tiff"],
    "bmp":  ["png", "jpg", "webp", "gif", "tiff"],
    "gif":  ["png", "jpg", "webp", "bmp", "tiff"],
    "tiff": ["png", "jpg", "webp", "bmp", "gif"],
    "tif":  ["png", "jpg", "webp", "bmp", "gif"],
    "ico":  ["png", "jpg"],
    "pdf":  ["txt"],
    "txt":  ["pdf"],
    "docx": ["txt"],
    "doc":  ["txt"],
    "md":   ["html"],
    "epub": ["pdf"],
    "csv":  ["json", "xml"],
    "json": ["csv", "xml"],
    "xml":  ["json"],
}

IMAGE_EXTS = {"png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff", "tif", "ico"}
PIL_FMT = {
    "jpg": "JPEG", "jpeg": "JPEG", "tif": "TIFF", "tiff": "TIFF",
    "png": "PNG", "webp": "WEBP", "bmp": "BMP", "gif": "GIF", "ico": "ICO",
}
EXT_FILETYPES = {
    "txt":  [("Text files", "*.txt"),   ("All files", "*.*")],
    "pdf":  [("PDF files", "*.pdf"),    ("All files", "*.*")],
    "html": [("HTML files", "*.html"),  ("All files", "*.*")],
    "json": [("JSON files", "*.json"),  ("All files", "*.*")],
    "csv":  [("CSV files", "*.csv"),    ("All files", "*.*")],
    "xml":  [("XML files", "*.xml"),    ("All files", "*.*")],
    "png":  [("PNG images", "*.png"),   ("All files", "*.*")],
    "jpg":  [("JPEG images", "*.jpg"),  ("All files", "*.*")],
    "jpeg": [("JPEG images", "*.jpeg"), ("All files", "*.*")],
    "webp": [("WebP images", "*.webp"), ("All files", "*.*")],
    "bmp":  [("BMP images", "*.bmp"),   ("All files", "*.*")],
    "gif":  [("GIF images", "*.gif"),   ("All files", "*.*")],
    "tiff": [("TIFF images", "*.tiff"), ("All files", "*.*")],
    "ico":  [("ICO images", "*.ico"),   ("All files", "*.*")],
    "zip":  [("ZIP files", "*.zip"),    ("All files", "*.*")],
}


# ── converters ────────────────────────────────────────────────────────────────

def convert_image(data, to_ext):
    from PIL import Image
    img = Image.open(io.BytesIO(data))
    if to_ext == "ico":
        img = img.convert("RGBA")
    elif to_ext in ("jpg", "jpeg", "bmp") and img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")
    out = io.BytesIO()
    img.save(out, format=PIL_FMT.get(to_ext, to_ext.upper()))
    return out.getvalue()


def pdf_to_txt(data):
    from pypdf import PdfReader
    pages = [p.extract_text() or "" for p in PdfReader(io.BytesIO(data)).pages]
    return ("\n\n--- Page Break ---\n\n".join(pages)).encode("utf-8")


def txt_to_pdf(data):
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas as rl_canvas
    text = data.decode("utf-8", errors="replace")
    out = io.BytesIO()
    c = rl_canvas.Canvas(out, pagesize=letter)
    _, h = letter
    y = h - 50
    for raw in text.split("\n"):
        line = raw
        while len(line) > 95:
            c.drawString(50, y, line[:95])
            line = line[95:]
            y -= 16
            if y < 50:
                c.showPage(); y = h - 50
        c.drawString(50, y, line)
        y -= 16
        if y < 50:
            c.showPage(); y = h - 50
    c.save()
    return out.getvalue()


def docx_to_txt(data):
    from docx import Document
    return "\n".join(p.text for p in Document(io.BytesIO(data)).paragraphs).encode("utf-8")


def md_to_html(data):
    import markdown as md_lib
    body = md_lib.markdown(data.decode("utf-8", errors="replace"), extensions=["extra"])
    return f'<!DOCTYPE html>\n<html>\n<head><meta charset="utf-8"></head>\n<body>\n{body}\n</body>\n</html>'.encode("utf-8")


def epub_to_pdf(data):
    import tempfile, warnings, ebooklib
    from ebooklib import epub
    from bs4 import BeautifulSoup
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from xml.sax.saxutils import escape

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".epub", delete=False) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            book = epub.read_epub(tmp_path)
        out = io.BytesIO()
        doc = SimpleDocTemplate(out, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=72)
        styles = getSampleStyleSheet()
        story = []
        for item in book.get_items_of_type(ebooklib.ITEM_DOCUMENT):
            soup = BeautifulSoup(item.get_content(), "html.parser")
            for tag in soup.find_all(["h1", "h2", "h3", "h4", "p", "li"]):
                text = escape(tag.get_text(separator=" ").strip())
                if not text:
                    continue
                if tag.name == "h1":
                    story.append(Paragraph(text, styles["Heading1"]))
                elif tag.name in ("h2", "h3", "h4"):
                    story.append(Paragraph(text, styles["Heading2"]))
                else:
                    story.append(Paragraph(text, styles["Normal"]))
                story.append(Spacer(1, 3))
        if not story:
            raise ValueError("No readable content found in EPUB")
        doc.build(story)
        return out.getvalue()
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


def csv_to_json(data):
    rows = list(csv.DictReader(io.StringIO(data.decode("utf-8", errors="replace"))))
    return json.dumps(rows, indent=2).encode("utf-8")


def json_to_csv(data):
    rows = json.loads(data.decode("utf-8"))
    if not rows or not isinstance(rows, list):
        raise ValueError("JSON must be an array of objects")
    keys = list(rows[0].keys())
    out = io.StringIO()
    w = csv.DictWriter(out, fieldnames=keys, extrasaction="ignore")
    w.writeheader(); w.writerows(rows)
    return out.getvalue().encode("utf-8")


def _to_xml_elem(parent, data):
    if isinstance(data, dict):
        for k, v in data.items():
            tag = "".join(c if c.isalnum() or c == "_" else "_" for c in k)
            if not tag or tag[0].isdigit():
                tag = "field_" + tag
            _to_xml_elem(ET.SubElement(parent, tag), v)
    elif isinstance(data, list):
        for item in data:
            _to_xml_elem(ET.SubElement(parent, "item"), item)
    else:
        parent.text = str(data) if data is not None else ""


def csv_to_xml(data):
    rows = list(csv.DictReader(io.StringIO(data.decode("utf-8", errors="replace"))))
    root = ET.Element("data")
    for row in rows:
        _to_xml_elem(ET.SubElement(root, "row"), row)
    return minidom.parseString(ET.tostring(root, encoding="unicode")).toprettyxml(indent="  ").encode("utf-8")


def json_to_xml(data):
    obj = json.loads(data.decode("utf-8"))
    root = ET.Element("root")
    _to_xml_elem(root, obj)
    return minidom.parseString(ET.tostring(root, encoding="unicode")).toprettyxml(indent="  ").encode("utf-8")


def xml_to_json(data):
    def elem_dict(elem):
        children = list(elem)
        if not children:
            return elem.text.strip() if elem.text and elem.text.strip() else None
        d = {}
        for child in children:
            val = elem_dict(child)
            if child.tag in d:
                if not isinstance(d[child.tag], list):
                    d[child.tag] = [d[child.tag]]
                d[child.tag].append(val)
            else:
                d[child.tag] = val
        return d
    root = ET.fromstring(data.decode("utf-8"))
    return json.dumps({root.tag: elem_dict(root)}, indent=2).encode("utf-8")


CONVERTERS = {
    ("pdf",  "txt"):  (pdf_to_txt,  ".txt"),
    ("txt",  "pdf"):  (txt_to_pdf,  ".pdf"),
    ("docx", "txt"):  (docx_to_txt, ".txt"),
    ("doc",  "txt"):  (docx_to_txt, ".txt"),
    ("md",   "html"): (md_to_html,  ".html"),
    ("epub", "pdf"):  (epub_to_pdf, ".pdf"),
    ("csv",  "json"): (csv_to_json, ".json"),
    ("json", "csv"):  (json_to_csv, ".csv"),
    ("csv",  "xml"):  (csv_to_xml,  ".xml"),
    ("json", "xml"):  (json_to_xml, ".xml"),
    ("xml",  "json"): (xml_to_json, ".json"),
}


# ── UI ────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Converter")
        self.resizable(False, False)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        self._build_convert(nb)
        self._build_merge(nb)
        self._build_trim(nb)
        self._build_split(nb)

    # schedule UI update safely from any thread
    def _ui(self, fn):
        self.after(0, fn)

    def _set_status(self, lbl, text, color="black"):
        self._ui(lambda: lbl.config(text=text, foreground=color))

    # ── Convert ──────────────────────────────────────────────────────────────

    def _build_convert(self, nb):
        f = ttk.Frame(nb, padding=14)
        nb.add(f, text="Convert")

        ttk.Label(f, text="File:").grid(row=0, column=0, sticky="w")
        self._cv_path = tk.StringVar()
        ttk.Entry(f, textvariable=self._cv_path, width=46, state="readonly").grid(row=0, column=1, padx=6)
        ttk.Button(f, text="Browse…", command=self._cv_browse).grid(row=0, column=2)

        ttk.Label(f, text="Convert to:").grid(row=1, column=0, sticky="w", pady=10)
        self._cv_fmt = tk.StringVar()
        self._cv_combo = ttk.Combobox(f, textvariable=self._cv_fmt, state="readonly", width=12)
        self._cv_combo.grid(row=1, column=1, sticky="w", padx=6)

        self._cv_btn = ttk.Button(f, text="Convert", command=self._cv_go)
        self._cv_btn.grid(row=2, column=1, sticky="w", padx=6)

        self._cv_lbl = ttk.Label(f, text="", wraplength=500)
        self._cv_lbl.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def _cv_browse(self):
        p = filedialog.askopenfilename(parent=self)
        if not p:
            return
        self._cv_path.set(p)
        ext = Path(p).suffix.lstrip(".").lower()
        targets = CONVERSION_MAP.get(ext, [])
        self._cv_combo["values"] = targets
        if targets:
            self._cv_combo.current(0)
            self._cv_lbl.config(text="")
        else:
            self._cv_combo.set("")
            self._cv_lbl.config(text=f"No conversions available for .{ext}", foreground="gray")

    def _cv_go(self):
        path = self._cv_path.get()
        to_ext = self._cv_fmt.get()
        if not path or not to_ext:
            return
        from_ext = Path(path).suffix.lstrip(".").lower()
        stem = Path(path).stem

        if from_ext in IMAGE_EXTS and to_ext in IMAGE_EXTS:
            out_ext = to_ext
        else:
            key = (from_ext, to_ext)
            if key not in CONVERTERS:
                self._set_status(self._cv_lbl, f"No converter for {from_ext} → {to_ext}", "red")
                return
            out_ext = CONVERTERS[key][1].lstrip(".")

        save_path = filedialog.asksaveasfilename(
            parent=self,
            initialfile=f"{stem}.{out_ext}",
            initialdir=str(Path(path).parent),
            defaultextension=f".{out_ext}",
            filetypes=EXT_FILETYPES.get(out_ext, [("All files", "*.*")]),
        )
        if not save_path:
            return

        self._cv_btn.state(["disabled"])
        self._set_status(self._cv_lbl, "Converting…", "black")

        def work():
            try:
                data = Path(path).read_bytes()
                if from_ext in IMAGE_EXTS and to_ext in IMAGE_EXTS:
                    result = convert_image(data, to_ext)
                else:
                    fn, _ = CONVERTERS[(from_ext, to_ext)]
                    result = fn(data)
                Path(save_path).write_bytes(result)
                self._set_status(self._cv_lbl, f"Saved → {Path(save_path).name}", "green")
            except Exception as e:
                self._set_status(self._cv_lbl, f"Error: {e}", "red")
            finally:
                self._ui(lambda: self._cv_btn.state(["!disabled"]))

        threading.Thread(target=work, daemon=True).start()

    # ── PDF Merge ─────────────────────────────────────────────────────────────

    def _build_merge(self, nb):
        f = ttk.Frame(nb, padding=14)
        nb.add(f, text="PDF Merge")

        ttk.Label(f, text="PDFs to merge (top = first):").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        lf = ttk.Frame(f)
        lf.grid(row=1, column=0, columnspan=3, sticky="nsew")
        self._mg_lb = tk.Listbox(lf, height=6, width=58, selectmode=tk.SINGLE)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self._mg_lb.yview)
        self._mg_lb.config(yscrollcommand=sb.set)
        self._mg_lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._mg_paths = []

        bf = ttk.Frame(f)
        bf.grid(row=2, column=0, columnspan=3, sticky="w", pady=8)
        for txt, cmd in [
            ("Add Files…", self._mg_add),
            ("Remove",     self._mg_remove),
            ("↑",          lambda: self._mg_move(-1)),
            ("↓",          lambda: self._mg_move(1)),
        ]:
            ttk.Button(bf, text=txt, command=cmd).pack(side="left", padx=(0, 4))

        self._mg_btn = ttk.Button(f, text="Merge", command=self._mg_go)
        self._mg_btn.grid(row=3, column=0, sticky="w")

        self._mg_lbl = ttk.Label(f, text="", wraplength=500)
        self._mg_lbl.grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def _mg_add(self):
        paths = filedialog.askopenfilenames(parent=self, filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        for p in paths:
            self._mg_paths.append(p)
            self._mg_lb.insert(tk.END, Path(p).name)

    def _mg_remove(self):
        sel = self._mg_lb.curselection()
        if not sel:
            return
        i = sel[0]
        self._mg_lb.delete(i)
        self._mg_paths.pop(i)

    def _mg_move(self, d):
        sel = self._mg_lb.curselection()
        if not sel:
            return
        i, j = sel[0], sel[0] + d
        if j < 0 or j >= len(self._mg_paths):
            return
        self._mg_paths[i], self._mg_paths[j] = self._mg_paths[j], self._mg_paths[i]
        items = list(self._mg_lb.get(0, tk.END))
        items[i], items[j] = items[j], items[i]
        self._mg_lb.delete(0, tk.END)
        for item in items:
            self._mg_lb.insert(tk.END, item)
        self._mg_lb.select_set(j)

    def _mg_go(self):
        if len(self._mg_paths) < 2:
            self._set_status(self._mg_lbl, "Add at least 2 PDF files.", "red")
            return
        save_path = filedialog.asksaveasfilename(
            parent=self,
            initialfile="merged.pdf",
            initialdir=str(Path(self._mg_paths[0]).parent),
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not save_path:
            return

        self._mg_btn.state(["disabled"])
        self._set_status(self._mg_lbl, "Merging…", "black")
        paths = list(self._mg_paths)

        def work():
            try:
                from pypdf import PdfReader, PdfWriter
                writer = PdfWriter()
                for p in paths:
                    for page in PdfReader(p).pages:
                        writer.add_page(page)
                with open(save_path, "wb") as fh:
                    writer.write(fh)
                self._set_status(self._mg_lbl, f"Saved → {Path(save_path).name}", "green")
            except Exception as e:
                self._set_status(self._mg_lbl, f"Error: {e}", "red")
            finally:
                self._ui(lambda: self._mg_btn.state(["!disabled"]))

        threading.Thread(target=work, daemon=True).start()

    # ── PDF Trim ──────────────────────────────────────────────────────────────

    def _build_trim(self, nb):
        f = ttk.Frame(nb, padding=14)
        nb.add(f, text="PDF Trim")

        ttk.Label(f, text="File:").grid(row=0, column=0, sticky="w")
        self._tr_path = tk.StringVar()
        ttk.Entry(f, textvariable=self._tr_path, width=46, state="readonly").grid(row=0, column=1, padx=6)
        ttk.Button(f, text="Browse…", command=self._tr_browse).grid(row=0, column=2)

        self._tr_info = ttk.Label(f, text="", foreground="gray")
        self._tr_info.grid(row=1, column=1, sticky="w", padx=6, pady=2)

        ttk.Label(f, text="Pages:").grid(row=2, column=0, sticky="w", pady=10)
        pf = ttk.Frame(f)
        pf.grid(row=2, column=1, sticky="w", padx=6)
        ttk.Label(pf, text="from").pack(side="left")
        self._tr_start = ttk.Spinbox(pf, from_=1, to=9999, width=7)
        self._tr_start.set(1)
        self._tr_start.pack(side="left", padx=4)
        ttk.Label(pf, text="to").pack(side="left")
        self._tr_end = ttk.Spinbox(pf, from_=1, to=9999, width=7)
        self._tr_end.pack(side="left", padx=4)

        self._tr_btn = ttk.Button(f, text="Trim", command=self._tr_go)
        self._tr_btn.grid(row=3, column=1, sticky="w", padx=6)

        self._tr_lbl = ttk.Label(f, text="", wraplength=500)
        self._tr_lbl.grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def _tr_browse(self):
        p = filedialog.askopenfilename(parent=self, filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not p:
            return
        self._tr_path.set(p)
        try:
            from pypdf import PdfReader
            n = len(PdfReader(p).pages)
            self._tr_info.config(text=f"{n} pages")
            self._tr_start.config(to=n)
            self._tr_end.config(to=n)
            self._tr_end.set(n)
        except Exception as e:
            self._tr_info.config(text=f"Error: {e}")

    def _tr_go(self):
        path = self._tr_path.get()
        if not path:
            return
        try:
            start, end = int(self._tr_start.get()), int(self._tr_end.get())
        except ValueError:
            self._set_status(self._tr_lbl, "Enter valid page numbers.", "red")
            return

        save_path = filedialog.asksaveasfilename(
            parent=self,
            initialfile=f"{Path(path).stem}_pages_{start}-{end}.pdf",
            initialdir=str(Path(path).parent),
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not save_path:
            return

        self._tr_btn.state(["disabled"])
        self._set_status(self._tr_lbl, "Trimming…", "black")

        def work():
            try:
                from pypdf import PdfReader, PdfWriter
                reader = PdfReader(path)
                total = len(reader.pages)
                if start < 1 or end > total or start > end:
                    raise ValueError(f"Invalid range {start}–{end}. PDF has {total} pages.")
                writer = PdfWriter()
                for i in range(start - 1, end):
                    writer.add_page(reader.pages[i])
                with open(save_path, "wb") as fh:
                    writer.write(fh)
                self._set_status(self._tr_lbl, f"Saved → {Path(save_path).name}", "green")
            except Exception as e:
                self._set_status(self._tr_lbl, f"Error: {e}", "red")
            finally:
                self._ui(lambda: self._tr_btn.state(["!disabled"]))

        threading.Thread(target=work, daemon=True).start()

    # ── PDF Split ─────────────────────────────────────────────────────────────

    def _build_split(self, nb):
        f = ttk.Frame(nb, padding=14)
        nb.add(f, text="PDF Split")

        ttk.Label(f, text="File:").grid(row=0, column=0, sticky="w")
        self._sp_path = tk.StringVar()
        ttk.Entry(f, textvariable=self._sp_path, width=46, state="readonly").grid(row=0, column=1, padx=6)
        ttk.Button(f, text="Browse…", command=self._sp_browse).grid(row=0, column=2)

        self._sp_info = ttk.Label(f, text="", foreground="gray")
        self._sp_info.grid(row=1, column=1, sticky="w", padx=6, pady=2)

        ttk.Label(f, text="Pages/chunk:").grid(row=2, column=0, sticky="w", pady=10)
        self._sp_n = ttk.Spinbox(f, from_=1, to=9999, width=7)
        self._sp_n.set(1)
        self._sp_n.grid(row=2, column=1, sticky="w", padx=6)

        self._sp_btn = ttk.Button(f, text="Split → .zip", command=self._sp_go)
        self._sp_btn.grid(row=3, column=1, sticky="w", padx=6)

        self._sp_lbl = ttk.Label(f, text="", wraplength=500)
        self._sp_lbl.grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))

    def _sp_browse(self):
        p = filedialog.askopenfilename(parent=self, filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not p:
            return
        self._sp_path.set(p)
        try:
            from pypdf import PdfReader
            n = len(PdfReader(p).pages)
            self._sp_info.config(text=f"{n} pages")
            self._sp_n.config(to=n)
        except Exception as e:
            self._sp_info.config(text=f"Error: {e}")

    def _sp_go(self):
        path = self._sp_path.get()
        if not path:
            return
        try:
            n = max(1, int(self._sp_n.get()))
        except ValueError:
            self._set_status(self._sp_lbl, "Enter a valid number.", "red")
            return

        save_path = filedialog.asksaveasfilename(
            parent=self,
            initialfile=f"{Path(path).stem}_split.zip",
            initialdir=str(Path(path).parent),
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")],
        )
        if not save_path:
            return

        self._sp_btn.state(["disabled"])
        self._set_status(self._sp_lbl, "Splitting…", "black")

        def work():
            try:
                from pypdf import PdfReader, PdfWriter
                stem = Path(path).stem
                reader = PdfReader(path)
                total = len(reader.pages)
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    i = 0
                    while i < total:
                        chunk_end = min(i + n, total)
                        writer = PdfWriter()
                        for j in range(i, chunk_end):
                            writer.add_page(reader.pages[j])
                        buf = io.BytesIO()
                        writer.write(buf)
                        name = f"{stem}_page_{i+1}.pdf" if n == 1 else f"{stem}_pages_{i+1}-{chunk_end}.pdf"
                        zf.writestr(name, buf.getvalue())
                        i += n
                Path(save_path).write_bytes(zip_buf.getvalue())
                self._set_status(self._sp_lbl, f"Saved → {Path(save_path).name}", "green")
            except Exception as e:
                self._set_status(self._sp_lbl, f"Error: {e}", "red")
            finally:
                self._ui(lambda: self._sp_btn.state(["!disabled"]))

        threading.Thread(target=work, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
