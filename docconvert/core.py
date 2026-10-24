"""Conversion logic.

Every reader turns a file into a list of Block(kind, text), where kind is
"h1".."h6" or "p". Every writer turns that list back into a file. That keeps
the matrix small: adding a format means one reader and one writer.

When LibreOffice (``soffice``) is on PATH, docx<->pdf goes through it for
layout fidelity. Otherwise the pure-Python path is used.
"""

from __future__ import annotations

import html
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

SUPPORTED = ("docx", "pdf", "md", "txt", "html")


@dataclass
class Block:
    kind: str  # "p" or "h1".."h6"
    text: str


def _ext(path: Path) -> str:
    return path.suffix.lower().lstrip(".")


# ---------------------------------------------------------------- readers


def read_txt(path: Path) -> list[Block]:
    raw = path.read_text(encoding="utf-8", errors="replace")
    return [Block("p", chunk.strip()) for chunk in re.split(r"\n\s*\n", raw) if chunk.strip()]


def read_md(path: Path) -> list[Block]:
    blocks: list[Block] = []
    raw = path.read_text(encoding="utf-8", errors="replace")
    for chunk in re.split(r"\n\s*\n", raw):
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.match(r"^(#{1,6})\s+(.*)$", chunk, re.S)
        if m:
            blocks.append(Block(f"h{len(m.group(1))}", m.group(2).strip()))
        else:
            blocks.append(Block("p", chunk))
    return blocks


def read_html(path: Path) -> list[Block]:
    from html.parser import HTMLParser

    class P(HTMLParser):
        def __init__(self) -> None:
            super().__init__()
            self.blocks: list[Block] = []
            self.kind: str | None = None
            self.buf: list[str] = []

        def handle_starttag(self, tag, attrs):
            if tag in ("p", "h1", "h2", "h3", "h4", "h5", "h6", "li"):
                self.kind = "p" if tag == "li" else tag
                self.buf = []
            elif tag == "br" and self.kind:
                self.buf.append("\n")

        def handle_endtag(self, tag):
            if self.kind and tag in ("p", "h1", "h2", "h3", "h4", "h5", "h6", "li"):
                text = "".join(self.buf).strip()
                if text:
                    self.blocks.append(Block(self.kind, text))
                self.kind = None

        def handle_data(self, data):
            if self.kind:
                self.buf.append(data)

    parser = P()
    parser.feed(path.read_text(encoding="utf-8", errors="replace"))
    return parser.blocks


def read_docx(path: Path) -> list[Block]:
    import docx

    blocks: list[Block] = []
    for para in docx.Document(str(path)).paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = (para.style.name or "").lower()
        m = re.match(r"heading (\d)", style)
        kind = f"h{m.group(1)}" if m else ("h1" if style == "title" else "p")
        blocks.append(Block(kind, text))
    return blocks


def read_pdf(path: Path) -> list[Block]:
    from pypdf import PdfReader

    text = "\n".join((page.extract_text() or "") for page in PdfReader(str(path)).pages)
    return [Block("p", chunk.strip()) for chunk in re.split(r"\n\s*\n", text) if chunk.strip()]


# ---------------------------------------------------------------- writers


def write_txt(blocks: list[Block], path: Path) -> None:
    path.write_text("\n\n".join(b.text for b in blocks) + "\n", encoding="utf-8")


def write_md(blocks: list[Block], path: Path) -> None:
    out = []
    for b in blocks:
        if b.kind.startswith("h"):
            out.append("#" * int(b.kind[1]) + " " + b.text)
        else:
            out.append(b.text)
    path.write_text("\n\n".join(out) + "\n", encoding="utf-8")


def write_html(blocks: list[Block], path: Path) -> None:
    body = "\n".join(
        f"<{b.kind}>{html.escape(b.text).replace(chr(10), '<br>')}</{b.kind}>" for b in blocks
    )
    path.write_text(
        "<!doctype html>\n<html><head><meta charset=\"utf-8\">"
        f"<title>{html.escape(path.stem)}</title></head>\n<body>\n{body}\n</body></html>\n",
        encoding="utf-8",
    )


def write_docx(blocks: list[Block], path: Path) -> None:
    import docx

    doc = docx.Document()
    for b in blocks:
        if b.kind.startswith("h"):
            doc.add_heading(b.text, level=int(b.kind[1]))
        else:
            doc.add_paragraph(b.text)
    doc.save(str(path))


def write_pdf(blocks: list[Block], path: Path) -> None:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

    styles = getSampleStyleSheet()
    story = []
    for b in blocks:
        style = styles["Heading" + b.kind[1]] if b.kind.startswith("h") else styles["BodyText"]
        story.append(Paragraph(html.escape(b.text).replace("\n", "<br/>"), style))
        story.append(Spacer(1, 6))
    SimpleDocTemplate(str(path), pagesize=A4, title=path.stem).build(story)


READERS = {"txt": read_txt, "md": read_md, "html": read_html, "docx": read_docx, "pdf": read_pdf}
WRITERS = {"txt": write_txt, "md": write_md, "html": write_html, "docx": write_docx, "pdf": write_pdf}

# ---------------------------------------------------------------- libreoffice


def _soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    return str(mac) if mac.exists() else None


def _convert_with_soffice(src: Path, dst: Path, exe: str) -> None:
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(
            [exe, "--headless", "--convert-to", _ext(dst), "--outdir", tmp, str(src)],
            check=True,
            capture_output=True,
            timeout=120,
        )
        produced = Path(tmp) / (src.stem + "." + _ext(dst))
        if not produced.exists():
            raise RuntimeError("LibreOffice did not produce an output file")
        shutil.move(str(produced), str(dst))


# ---------------------------------------------------------------- public API


def convert(src: str | Path, dst: str | Path, *, use_soffice: bool = True) -> Path:
    """Convert ``src`` to ``dst``; the formats are taken from the extensions."""
    src, dst = Path(src), Path(dst)
    sfmt, dfmt = _ext(src), _ext(dst)
    if sfmt not in READERS:
        raise ValueError(f"unsupported input format: .{sfmt} (supported: {', '.join(SUPPORTED)})")
    if dfmt not in WRITERS:
        raise ValueError(f"unsupported output format: .{dfmt} (supported: {', '.join(SUPPORTED)})")
    if not src.is_file():
        raise FileNotFoundError(src)
    if sfmt == dfmt:
        raise ValueError("input and output formats are the same")

    dst.parent.mkdir(parents=True, exist_ok=True)

    if use_soffice and sfmt == "docx" and dfmt == "pdf":
        exe = _soffice()
        if exe:
            _convert_with_soffice(src, dst, exe)
            return dst

    WRITERS[dfmt](READERS[sfmt](src), dst)
    return dst
