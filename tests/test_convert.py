import itertools
from pathlib import Path

import pytest

from docconvert import SUPPORTED, convert
from docconvert.core import READERS

SAMPLE_MD = "# Title\n\nFirst paragraph with some text.\n\n## Section\n\nSecond paragraph.\n"


@pytest.fixture
def sample_md(tmp_path: Path) -> Path:
    p = tmp_path / "sample.md"
    p.write_text(SAMPLE_MD, encoding="utf-8")
    return p


@pytest.mark.parametrize("src_fmt,dst_fmt", [p for p in itertools.permutations(SUPPORTED, 2)])
def test_every_pair_produces_text(tmp_path: Path, sample_md: Path, src_fmt: str, dst_fmt: str):
    src = sample_md if src_fmt == "md" else convert(sample_md, tmp_path / f"src.{src_fmt}", use_soffice=False)
    dst = convert(src, tmp_path / f"out.{dst_fmt}", use_soffice=False)
    assert dst.is_file() and dst.stat().st_size > 0
    text = " ".join(b.text for b in READERS[dst_fmt](dst))
    assert "First paragraph" in text
    assert "Second paragraph" in text


def test_headings_survive_md_docx_md(tmp_path: Path, sample_md: Path):
    docx = convert(sample_md, tmp_path / "a.docx")
    back = convert(docx, tmp_path / "b.md")
    assert back.read_text().startswith("# Title")
    assert "## Section" in back.read_text()


def test_rejects_unknown_format(tmp_path: Path, sample_md: Path):
    with pytest.raises(ValueError):
        convert(sample_md, tmp_path / "x.xyz")


def test_missing_input(tmp_path: Path):
    with pytest.raises(FileNotFoundError):
        convert(tmp_path / "nope.md", tmp_path / "x.pdf")


def test_cli_batch(tmp_path: Path, sample_md: Path):
    from docconvert.cli import main

    assert main([str(sample_md), "--to", "txt", "--no-soffice"]) == 0
    assert (tmp_path / "sample.txt").is_file()
