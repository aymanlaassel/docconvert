# docconvert

A small tool to switch documents between formats. No fancy stuff.

Supported formats: `docx`, `pdf`, `md`, `txt`, `html` (any direction).

## Getting started

### 1. Install Python

You need Python 3.10 or newer. Check with:

```sh
python3 --version
```

If it's missing, download it from https://www.python.org/downloads/ and install it.

### 2. Get the code

```sh
git clone https://github.com/aymanlaassel/docconvert.git
cd docconvert
```

(No git? Click the green **Code** button on GitHub, choose **Download ZIP**,
unzip it, then `cd` into the folder.)

### 3. Install the tool

```sh
pip install .
```

This installs two commands: `docconvert` and `docconvert-ui`.

If `pip` isn't found, try `pip3 install .` or `python3 -m pip install .`.

### 4. Run it

**Easiest: the web UI**

```sh
docconvert-ui
```

A page opens in your browser. Drop a file in, pick the format, click
**Convert**. The converted file lands in your Downloads folder. Press
`Ctrl+C` in the terminal to stop it.

**Or: the command line**

```sh
docconvert ~/Downloads/report.docx -o report.pdf   # explicit output file
docconvert notes.md --to docx                      # same name, new extension
docconvert *.docx --to pdf                         # batch
```

Give the real path to your file. If the file is in another folder, either
`cd` there first or write the full path, like `~/Downloads/file.pdf`.

## Troubleshooting

- `command not found: docconvert` — your Python scripts folder isn't on PATH.
  Run it as `python3 -m docconvert ...` or `python3 -m docconvert.ui` instead.
- `file not found: something.pdf` — the file isn't in the folder you're in.
  Check the exact name (spaces and capitals matter) or use the full path.
- Permission errors on `pip install .` — try `pip install --user .`.

## What it does and doesn't do

Files are read into a flat list of headings and paragraphs, then written back
out in the target format. Formatting beyond headings and paragraphs (tables,
images, fonts, columns) is not preserved.

If [LibreOffice](https://www.libreoffice.org/) is installed, `docx -> pdf` goes
through it for faithful layout. Pass `--no-soffice` to force the pure-Python path.

## Develop

```sh
python3 -m venv .venv && . .venv/bin/activate
pip install -e ".[dev]"
pytest
```

## License

MIT
