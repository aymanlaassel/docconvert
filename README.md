# docconvert

A small command line tool to switch documents between formats. No fancy stuff.

Supported formats: `docx`, `pdf`, `md`, `txt`, `html` (any direction).

## Install

```sh
pip install .
```

## Use

```sh
docconvert report.docx -o report.pdf      # explicit output file
docconvert notes.md --to docx             # same name, new extension
docconvert *.docx --to pdf                # batch
```

Or without installing: `python -m docconvert report.docx -o report.pdf`.

## How it works

Files are read into a flat list of headings and paragraphs, then written back
out in the target format. Formatting beyond headings and paragraphs (tables,
images, fonts) is not preserved.

If [LibreOffice](https://www.libreoffice.org/) is installed, `docx -> pdf` goes
through it for faithful layout. Pass `--no-soffice` to force the pure-Python path.

## Develop

```sh
python -m venv .venv && . .venv/bin/activate
pip install -e ".[dev]"
pytest
```

## License

MIT
