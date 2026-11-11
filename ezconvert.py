"""EZConvert: convert images, PDF, MD, DOCX, TXT from the right-click menu or a drop window.
Deps: pip install pillow pymupdf pymupdf4llm pypandoc_binary docx2pdf python-docx tkinterdnd2
DOCX->PDF uses Microsoft Word; MD->PDF uses Microsoft Edge (both installed).
Heavy libraries are imported only by the conversion that needs them, so each click starts fast.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

if sys.stdout is None:  # pythonw has no console; some libs print progress
    sys.stdout = sys.stderr = open(os.devnull, "w")

IMG = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".ico", ".tiff"}
TXT = {".txt", ".md"}

PDF_CSS = """
@page{margin:14mm 14mm}
body{font-family:"Segoe UI",Arial,sans-serif;font-size:10.5pt;line-height:1.45;color:#1f2328;max-width:none;margin:0;padding:0}
header#title-block-header{display:none}
h1,h2{border-bottom:1px solid #d8dee4;padding-bottom:.3em}
h1{font-size:2em} h2{font-size:1.5em} h3{font-size:1.2em}
h1,h2,h3{margin:1.1em 0 .5em;font-weight:600;page-break-after:avoid}
code{font-family:Consolas,monospace;background:#eff1f3;padding:.1em .35em;border-radius:4px;font-size:90%}
pre{background:#f6f8fa;padding:8px 12px;border-radius:6px;overflow:auto;page-break-inside:avoid}
pre code{background:none;padding:0}
blockquote{margin:0;padding:0 1em;color:#59636e;border-left:4px solid #d1d9e0}
table{border-collapse:collapse;width:100%;margin:1em 0;page-break-inside:auto}
tr{page-break-inside:avoid}
th,td{border:1px solid #d1d9e0;padding:4px 12px;text-align:left}
th{background:#f6f8fa;font-weight:600;text-align:center}
tr:nth-child(even) td{background:#fafbfc}
img{max-width:100%}
a{color:#0969da;text-decoration:none}
"""

CSS = """
body{font-family:sans-serif;font-size:10.5pt}
h1{font-size:22pt} h2{font-size:17pt} h3{font-size:14pt}
code,pre{font-family:monospace;background-color:#f2f2f2}
pre{padding:6px}
blockquote{margin-left:14px;color:#555555}
table{border-collapse:collapse}
th,td{border:1px solid #888888;padding:3px}
th{background-color:#eeeeee}
img{max-width:100%}
"""


def targets_for(ext):
    if ext in IMG:
        return ["png", "jpg", "webp", "bmp", "ico", "pdf"]
    if ext == ".pdf":
        return ["docx", "md", "txt", "png"]
    if ext == ".md":
        return ["pdf", "docx", "html", "txt"]
    if ext == ".docx":
        return ["pdf", "md", "html", "txt"]
    if ext == ".txt":
        return ["pdf", "docx", "md", "html"]
    return []


def unique(path):
    """notes.pdf -> 'notes (1).pdf' if taken, so nothing is ever overwritten."""
    path = Path(path)
    n = 1
    cand = path
    while cand.exists():
        cand = path.with_name(f"{path.stem} ({n}){path.suffix}")
        n += 1
    return cand


def find_edge():
    for base in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"),
                 os.environ.get("LOCALAPPDATA")):
        if base:
            p = Path(base) / "Microsoft" / "Edge" / "Application" / "msedge.exe"
            if p.exists():
                return str(p)
    found = shutil.which("msedge")
    if found:
        return found
    raise RuntimeError("Microsoft Edge not found (needed for MD -> PDF)")


def polish_docx(path):
    """Give pandoc's DOCX visible table borders, shaded headers and dark headings."""
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import RGBColor
    d = Document(str(path))
    for n in range(1, 7):
        try:
            d.styles[f"Heading {n}"].font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
        except KeyError:
            pass
    for t in d.tables:
        pr = t._tbl.tblPr
        b = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement(f"w:{side}")
            for k, v in (("val", "single"), ("sz", "4"), ("space", "0"), ("color", "BFC6CE")):
                e.set(qn(f"w:{k}"), v)
            b.append(e)
        pr.append(b)
        m = OxmlElement("w:tblCellMar")
        for side in ("top", "bottom"):
            e = OxmlElement(f"w:{side}")
            e.set(qn("w:w"), "40")
            e.set(qn("w:type"), "dxa")
            m.append(e)
        for side in ("left", "right"):
            e = OxmlElement(f"w:{side}")
            e.set(qn("w:w"), "100")
            e.set(qn("w:type"), "dxa")
            m.append(e)
        pr.append(m)
        for c in t.rows[0].cells:
            sh = OxmlElement("w:shd")
            sh.set(qn("w:val"), "clear")
            sh.set(qn("w:fill"), "F0F3F6")
            c._tc.get_or_add_tcPr().append(sh)
            for p in c.paragraphs:
                for r in p.runs:
                    r.bold = True
    d.save(str(path))


def tidy_md(text):
    """Clean extraction artefacts: stray spaces around bold, loose list gaps, trailing spaces."""
    text = re.sub(r"[ \t]+$", "", text, flags=re.M)
    text = re.sub(r"\*\* ?\*\*", "", text)                      # adjacent bold runs
    text = re.sub(r"(\*\*) +([,.;:!?)\]])", r"\1\2", text)      # '**bold** ,' -> '**bold**,'
    text = re.sub(r"(\*\*|_) +(?=[⌊⌈])", r"\1", text)
    text = re.sub(r"^((?:[-*] |\d+\. ).*)\n\n(?=(?:[-*] |\d+\. ))", r"\1\n", text, flags=re.M)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def unescape_md(text):
    return text.replace("\\[", "[").replace("\\]", "]")


def to_styled_html(src, frm, out, *extra):
    with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
        css = Path(td) / "style.css"
        css.write_text(PDF_CSS.replace("@page{margin:14mm 14mm}", "body{max-width:900px;margin:2em auto;padding:0 1em}"),
                       encoding="utf-8")
        pandoc(src, frm, "html", out, "--standalone", "--embed-resources", f"--css={css}",
               "--metadata", f"title={src.stem}", "--metadata", "pagetitle=" + src.stem, *extra)


def pandoc(src, frm, to, out, *extra):
    import pypandoc
    pypandoc.convert_file(str(src), to, format=frm, outputfile=str(out),
                          extra_args=["--wrap=none", *extra])


def html_fragment(src):
    if src.suffix == ".md":
        import pypandoc
        return pypandoc.convert_file(str(src), "html", format="gfm")
    text = src.read_text(encoding="utf-8", errors="replace")
    return "".join(f"<p>{l}</p>" for l in text.splitlines())


def html_to_pdf(html, base, out):
    import pymupdf
    w = pymupdf.DocumentWriter(str(out))
    story = pymupdf.Story(html=html, user_css=CSS, archive=pymupdf.Archive(str(base)))
    page = pymupdf.paper_rect("a4")
    more = 1
    while more:
        dev = w.begin_page(page)
        more, _ = story.place(page + (40, 40, -40, -40))
        story.draw(dev)
        w.end_page()
    w.close()


def pdf_to_md(src):
    import pymupdf4llm
    return tidy_md(pymupdf4llm.to_markdown(str(src)))


def convert(src, fmt, outdir):
    src = Path(src)
    ext = src.suffix.lower()
    outdir = Path(outdir).resolve()
    out = unique(outdir / f"{src.stem}.{fmt}")
    if ext in IMG:
        from PIL import Image
        im = Image.open(src)
        if fmt in ("jpg", "bmp", "pdf"):
            im = im.convert("RGB")
        im.save(out, {"jpg": "JPEG", "pdf": "PDF", "ico": "ICO"}.get(fmt, fmt.upper()))
    elif ext == ".pdf":
        if fmt == "docx":
            # PDF -> Markdown -> DOCX: clean, editable structure
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
                tmp = Path(td) / f"{src.stem}.md"
                tmp.write_text(pdf_to_md(src), encoding="utf-8")
                pandoc(tmp, "gfm", "docx", out)
            polish_docx(out)
        elif fmt == "md":
            out.write_text(pdf_to_md(src), encoding="utf-8")
        elif fmt == "txt":
            import pymupdf
            out.write_text("\n".join(p.get_text() for p in pymupdf.open(src)), encoding="utf-8")
        else:
            import pymupdf
            for i, p in enumerate(pymupdf.open(src), 1):
                out = unique(outdir / f"{src.stem}_{i}.png")
                p.get_pixmap(dpi=150).save(out)
    elif ext == ".docx":
        if fmt == "pdf":
            from docx2pdf import convert as word_pdf
            word_pdf(str(src), str(out))
        elif fmt == "md":
            media = unique(outdir / f"{out.stem}_media")
            pandoc(src, "docx", "gfm", out, f"--extract-media={media}")
            out.write_text(unescape_md(out.read_text(encoding="utf-8")), encoding="utf-8")
        elif fmt == "html":
            to_styled_html(src, "docx", out)
        else:
            pandoc(src, "docx", "plain", out)
    elif ext == ".md":
        if fmt == "pdf":  # md -> styled html -> Edge print: GitHub-like look
            with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as td:
                css = Path(td) / "style.css"
                css.write_text(PDF_CSS, encoding="utf-8")
                page = Path(td) / "page.html"
                pandoc(src, "gfm", "html", page, "--standalone", "--embed-resources",
                       f"--css={css}", f"--resource-path={src.parent}",
                       "--metadata", f"title={src.stem}", "--metadata", "pagetitle=doc")
                subprocess.run([find_edge(), "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
                                f"--user-data-dir={Path(td) / 'profile'}",
                                f"--print-to-pdf={out}", page.as_uri()],
                               check=True, timeout=120, capture_output=True,
                               creationflags=0x08000000)
                last = -1  # Edge returns before the file is fully written
                for _ in range(120):
                    size = out.stat().st_size if out.exists() else -1
                    if size > 0 and size == last:
                        break
                    last = size
                    time.sleep(0.25)
                else:
                    raise RuntimeError("Edge did not produce a PDF")
        elif fmt == "docx":
            pandoc(src, "gfm", "docx", out, f"--resource-path={src.parent}")
            polish_docx(out)
        elif fmt == "html":
            to_styled_html(src, "gfm", out, f"--resource-path={src.parent}")
        else:
            pandoc(src, "gfm", "plain", out)
    else:  # .txt
        text = src.read_text(encoding="utf-8", errors="replace")
        if fmt == "md":
            out.write_text(text, encoding="utf-8")
        elif fmt == "html":
            out.write_text(html_fragment(src), encoding="utf-8")
        elif fmt == "docx":
            from docx import Document
            d = Document()
            for line in text.splitlines():
                d.add_paragraph(line)
            d.save(out)
        elif fmt == "pdf":
            html_to_pdf(html_fragment(src), src.parent, out)
    return out


class Lock:
    """System-wide lock so several right-click conversions run one after another
    (Explorer starts one process per selected file; Word/Edge dislike parallel use)."""
    PATH = Path(tempfile.gettempdir()) / "ezconvert.lock"

    def __enter__(self):
        import msvcrt
        self.f = open(self.PATH, "a+")
        self.f.seek(0)
        while True:
            try:
                msvcrt.locking(self.f.fileno(), msvcrt.LK_NBLCK, 1)
                return self
            except OSError:
                time.sleep(0.2)

    def __exit__(self, *exc):
        import msvcrt
        self.f.seek(0)
        msvcrt.locking(self.f.fileno(), msvcrt.LK_UNLCK, 1)
        self.f.close()


def run_cli(fmt, files):
    """Right-click mode: small toast bottom-right showing progress, then Done or the error."""
    root = tk.Tk()
    root.overrideredirect(True)
    root.attributes("-topmost", True)
    root.configure(bg="#1f2328")
    label = tk.Label(root, bg="#1f2328", fg="white", font=("Segoe UI", 10),
                     justify="left", padx=14, pady=10, wraplength=320)
    label.pack()

    def show(text, color="white"):
        label.config(text=text, fg=color)
        root.update_idletasks()
        w, h = root.winfo_reqwidth(), root.winfo_reqheight()
        root.geometry(f"+{root.winfo_screenwidth() - w - 16}+{root.winfo_screenheight() - h - 60}")

    names = ", ".join(Path(f).name for f in files)
    show(f"Waiting… {names} → {fmt.upper()}")
    result = {}

    def work():
        outs, errs = [], []
        with Lock():
            root.after(0, show, f"Converting {names} → {fmt.upper()}…")
            for f in files:
                try:
                    outs.append(convert(f, fmt, Path(f).parent).name)
                except Exception as e:
                    errs.append(f"{Path(f).name}: {e}")
        result.update(outs=outs, errs=errs)
        root.after(0, finish)

    def finish():
        if result["errs"]:
            show("Failed:\n" + "\n".join(result["errs"]) + "\n(click to close)", "#ff8182")
            label.bind("<Button-1>", lambda e: root.destroy())
            root.after(15000, root.destroy)
        else:
            show("✓ Saved " + ", ".join(result["outs"]), "#7ee787")
            root.after(2500, root.destroy)

    threading.Thread(target=work, daemon=True).start()
    root.mainloop()


def run_gui():
    from tkinter import filedialog
    from tkinterdnd2 import DND_FILES, TkinterDnD

    class App(TkinterDnD.Tk):
        def __init__(self):
            super().__init__()
            self.title("EZConvert")
            self.geometry("360x260")
            self.files = []
            self.outdir = None
            self.drop = tk.Label(self, text="Drop files here\n(or click to browse)",
                                 relief="groove", bd=3, height=5, cursor="hand2")
            self.drop.pack(fill="x", padx=12, pady=12)
            self.drop.bind("<Button-1>", lambda e: self.set_files(filedialog.askopenfilenames()))
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", lambda e: self.set_files(self.tk.splitlist(e.data)))
            self.bar = tk.Frame(self)
            self.bar.pack()
            self.status = tk.Label(self, text="", wraplength=330)
            self.status.pack(pady=8)
            tk.Button(self, text="Output folder: same as source (click to change)",
                      command=self.pick_out).pack()

        def pick_out(self):
            d = filedialog.askdirectory()
            if d:
                self.outdir = d

        def set_files(self, files):
            self.files = [f for f in files if f]
            for w in self.bar.winfo_children():
                w.destroy()
            if not self.files:
                return
            exts = {Path(f).suffix.lower() for f in self.files}
            fmts = set.intersection(*[set(targets_for(e)) for e in exts])
            self.drop.config(text=f"{len(self.files)} file(s) ready")
            if not fmts:
                self.status.config(text="No common target format for these files.")
            for fmt in sorted(fmts):
                tk.Button(self.bar, text=fmt.upper(), width=6,
                          command=lambda f=fmt: self.run(f)).pack(side="left", padx=3)

        def run(self, fmt):
            self.status.config(text="Converting…")
            self.update_idletasks()
            done, errs = 0, []
            for f in self.files:
                try:
                    convert(f, fmt, self.outdir or Path(f).parent)
                    done += 1
                except Exception as e:
                    errs.append(f"{Path(f).name}: {e}")
            self.status.config(text=f"Converted {done}/{len(self.files)}" +
                               ("\n" + "\n".join(errs) if errs else ""))

    App().mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 3 and sys.argv[1] == "--to":  # right-click mode: --to fmt file...
        run_cli(sys.argv[2], sys.argv[3:])
    else:
        run_gui()
