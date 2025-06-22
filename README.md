# click2pptx

This project provides a small utility that converts a Freeplane HTML
export containing an image map into a PPTX presentation. Every clickable
area in the HTML is recreated as a transparent rectangle that points to
the associated link.

## Installation

```bash
pip install .
```

The command above installs the `beautifulsoup4` and `python-pptx`
dependencies required by the script.

## Usage

```
click2pptx [-i SOURCE.html] [-o DESTINATION.pptx]
```

- `-i`, `--input`: source HTML file. If omitted, the script will pick the
  first `*.html` file found in the current directory.
- `-o`, `--output`: path of the PPTX file to create. If omitted, an
  `output` folder is created (if it does not already exist) and the file
  `mind_map_clickable_YYYYMMDD_HHMMSS.pptx` is written there.

Example:

```bash
click2pptx -i my_export.html -o presentation.pptx
```

The program reads the HTML file, extracts the clickable areas and the
image used, then generates an equivalent PPTX. Each active region of the
image is covered by an invisible rectangle with the hyperlink defined in
the Freeplane export.

## Development

Install the development dependencies:

```bash
pip install -e ".[dev]"
```

Initialize the pre-commit hooks:

```bash
pre-commit install
```

Run the tests with coverage enabled:

```bash
pytest --cov=click2pptx
```

## Standalone binary

You can build a standalone executable using PyInstaller and the
provided `click2pptx.spec` file:

```bash
pip install pyinstaller
pyinstaller click2pptx.spec
```

The resulting binary will be available in the `dist` folder.
