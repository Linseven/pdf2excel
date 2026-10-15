# Contributing to PDF2Excel

Thanks for your interest in contributing! / 感谢你的贡献意愿！

## Quick Start

```bash
git clone https://github.com/OWNER/pdf2excel.git
cd pdf2excel
python3 -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
python main.py            # launch GUI
python main.py --cli in.pdf -o out.xlsx   # CLI mode
```

## Development Guidelines

- `converter.py` is the pure conversion engine — keep it free of GUI dependencies.
- `main.py` holds the pywebview bridge (`Api` class) and CLI entry.
- `ui/` is plain HTML/CSS/JS with **no build step and no external CDN assets** — keep it offline-friendly.
- User-facing strings are in Simplified Chinese; errors must raise `ConversionError` with a friendly message.

## Tests

Always run both suites before submitting a PR:

```bash
python tests/run_tests.py     # engine end-to-end tests
python tests/gui_smoke.py     # real-window GUI smoke test (macOS; hidden window)
```

## Pull Requests

1. Fork & create a feature branch from `main`.
2. Keep PRs focused; one feature or fix per PR.
3. Describe **why** the change is needed and how it was tested.
