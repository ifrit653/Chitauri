"""Manuscrit OCR - small Tkinter client around Tesseract.

Select scanned pages (images), pick the language (French / Malagasy),
and export the recognised text to a Word (.docx) file plus a .txt copy.
"""
import os
import queue
import shutil
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pytesseract
from docx import Document
from PIL import Image, ImageFilter, ImageOps

IMAGE_TYPES = [("Images", "*.png *.jpg *.jpeg *.tif *.tiff *.bmp *.webp")]
DEFAULT_INSTALL = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# --------------------------------------------------------------------------
# Tesseract location (bundled next to the exe, or installed on the system)
# --------------------------------------------------------------------------
def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def setup_tesseract() -> bool:
    bundled = resource_path(os.path.join("tesseract", "tesseract.exe"))
    candidates = [bundled, shutil.which("tesseract"), DEFAULT_INSTALL]
    for cmd in candidates:
        if cmd and os.path.exists(cmd):
            pytesseract.pytesseract.tesseract_cmd = cmd
            tessdata = os.path.join(os.path.dirname(cmd), "tessdata")
            if os.path.isdir(tessdata):
                os.environ["TESSDATA_PREFIX"] = tessdata
            return True
    return False


def missing_languages(lang: str) -> list:
    tessdata = os.environ.get("TESSDATA_PREFIX")
    if not tessdata:
        return []
    return [c for c in lang.split("+")
            if not os.path.exists(os.path.join(tessdata, f"{c}.traineddata"))]


# --------------------------------------------------------------------------
# Image preprocessing (Pillow only, keeps the exe small)
# --------------------------------------------------------------------------
def otsu_threshold(gray: Image.Image) -> int:
    hist = gray.histogram()
    total = sum(hist)
    sum_all = sum(i * h for i, h in enumerate(hist))
    sum_b = w_b = 0
    best, thresh = 0.0, 127
    for t in range(256):
        w_b += hist[t]
        if w_b == 0:
            continue
        w_f = total - w_b
        if w_f == 0:
            break
        sum_b += t * hist[t]
        m_b = sum_b / w_b
        m_f = (sum_all - sum_b) / w_f
        var = w_b * w_f * (m_b - m_f) ** 2
        if var > best:
            best, thresh = var, t
    return thresh


def preprocess(img: Image.Image, binarize: bool) -> Image.Image:
    img = ImageOps.exif_transpose(img).convert("L")
    img = ImageOps.autocontrast(img, cutoff=1)
    if max(img.size) < 2000:  # small scans: upscale for better recognition
        img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    img = img.filter(ImageFilter.MedianFilter(3))
    if binarize:
        t = otsu_threshold(img)
        img = img.point(lambda p: 255 if p > t else 0)
    return img


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Manuscrit OCR")
        self.geometry("680x560")
        self.minsize(600, 480)
        self.files: list = []
        self.q: queue.Queue = queue.Queue()
        self.out_path = tk.StringVar()
        self.lang = tk.StringVar(value="fra+mlg")
        self.psm = tk.StringVar(value="6")
        self.binarize = tk.BooleanVar(value=True)
        self.keep_breaks = tk.BooleanVar(value=False)
        self._build()
        self.after(100, self._poll)
        if not setup_tesseract():
            messagebox.showerror(
                "Tesseract not found",
                "Tesseract was not found. Install it or bundle it in a "
                "'tesseract' folder next to the program.")

    def _build(self):
        pad = {"padx": 10, "pady": 4}

        top = ttk.Frame(self)
        top.pack(fill="both", expand=True, **pad)
        self.listbox = tk.Listbox(top, selectmode="extended", height=8)
        sb = ttk.Scrollbar(top, command=self.listbox.yview)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        btns = ttk.Frame(top)
        btns.pack(side="left", padx=(8, 0), anchor="n")
        ttk.Button(btns, text="Add images...", command=self.add_files).pack(fill="x")
        ttk.Button(btns, text="Remove", command=self.remove_selected).pack(fill="x", pady=4)
        ttk.Button(btns, text="Clear", command=self.clear).pack(fill="x")

        opts = ttk.LabelFrame(self, text="Options")
        opts.pack(fill="x", **pad)
        ttk.Label(opts, text="Language").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Combobox(opts, textvariable=self.lang, width=12,
                     values=["fra+mlg", "fra", "mlg", "eng"]).grid(row=0, column=1, padx=6)
        ttk.Label(opts, text="Layout (psm)").grid(row=0, column=2, sticky="w", padx=6)
        ttk.Combobox(opts, textvariable=self.psm, width=5, state="readonly",
                     values=["3", "4", "6", "11"]).grid(row=0, column=3, padx=6)
        ttk.Checkbutton(opts, text="Clean image (black & white)",
                        variable=self.binarize).grid(row=1, column=0, columnspan=2,
                                                     sticky="w", padx=6, pady=2)
        ttk.Checkbutton(opts, text="Keep line breaks",
                        variable=self.keep_breaks).grid(row=1, column=2, columnspan=2,
                                                        sticky="w", padx=6)

        out = ttk.Frame(self)
        out.pack(fill="x", **pad)
        ttk.Label(out, text="Output .docx").pack(side="left")
        ttk.Entry(out, textvariable=self.out_path).pack(side="left", fill="x",
                                                        expand=True, padx=6)
        ttk.Button(out, text="Browse...", command=self.choose_output).pack(side="left")

        self.start_btn = ttk.Button(self, text="Start OCR", command=self.start)
        self.start_btn.pack(**pad)
        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", **pad)
        self.log = tk.Text(self, height=8, state="disabled")
        self.log.pack(fill="both", expand=True, **pad)

    # -- file list ---------------------------------------------------------
    def add_files(self):
        for f in filedialog.askopenfilenames(filetypes=IMAGE_TYPES):
            if f not in self.files:
                self.files.append(f)
                self.listbox.insert("end", Path(f).name)
        if self.files and not self.out_path.get():
            self.out_path.set(str(Path(self.files[0]).with_name("ocr_result.docx")))

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]

    def clear(self):
        self.listbox.delete(0, "end")
        self.files.clear()

    def choose_output(self):
        p = filedialog.asksaveasfilename(defaultextension=".docx",
                                         filetypes=[("Word document", "*.docx")])
        if p:
            self.out_path.set(p)

    # -- run ---------------------------------------------------------------
    def start(self):
        if not self.files:
            messagebox.showinfo("No files", "Add at least one image.")
            return
        if not self.out_path.get():
            self.choose_output()
            if not self.out_path.get():
                return
        missing = missing_languages(self.lang.get())
        if missing:
            messagebox.showerror("Missing language data",
                                 "Missing traineddata for: " + ", ".join(missing))
            return
        self.start_btn.config(state="disabled")
        self.progress.config(value=0, maximum=len(self.files))
        args = (list(self.files), self.lang.get(), self.binarize.get(),
                self.psm.get(), self.keep_breaks.get(), self.out_path.get())
        threading.Thread(target=self.work, args=args, daemon=True).start()

    def work(self, files, lang, binarize, psm, keep_breaks, out):
        try:
            doc = Document()
            all_text = []
            for i, f in enumerate(files, 1):
                self.q.put(("log", f"[{i}/{len(files)}] {Path(f).name}"))
                with Image.open(f) as im:
                    img = preprocess(im, binarize)
                text = pytesseract.image_to_string(
                    img, lang=lang, config=f"--psm {psm}").strip()
                if len(files) > 1:
                    doc.add_heading(Path(f).name, level=2)
                    all_text.append(f"=== {Path(f).name} ===")
                all_text.append(text)
                for block in text.split("\n\n"):
                    block = block.strip()
                    if not block:
                        continue
                    if keep_breaks:
                        doc.add_paragraph(block)
                    else:
                        doc.add_paragraph(" ".join(block.split()))
                if i < len(files):
                    doc.add_page_break()
                self.q.put(("progress", i))
            doc.save(out)
            Path(out).with_suffix(".txt").write_text("\n\n".join(all_text),
                                                     encoding="utf-8")
            self.q.put(("done", out))
        except Exception as e:  # noqa: BLE001
            self.q.put(("error", str(e)))

    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self._log(val)
                elif kind == "progress":
                    self.progress.config(value=val)
                elif kind == "done":
                    self._log(f"Saved: {val}")
                    self.start_btn.config(state="normal")
                    messagebox.showinfo("Done", f"Saved to:\n{val}")
                elif kind == "error":
                    self._log(f"Error: {val}")
                    self.start_btn.config(state="normal")
                    messagebox.showerror("Error", val)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _log(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")


if __name__ == "__main__":
    App().mainloop()
