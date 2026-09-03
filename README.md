# pyhandlexl

Read and write raw cell values in Excel `.xlsx` files — safely and simply.

> **Status:** early development (0.1.0). API is not yet stable.

`pyhandlexl` is a thin layer over [openpyxl](https://openpyxl.readthedocs.io/) for
the common case: getting values in and out of a spreadsheet in a specific order,
without touching formatting. It focuses on **safe writes** (Excel files corrupt
easily) and a small, predictable API.

## Install

Not on PyPI yet. From source:

```bash
git clone https://github.com/LewyAmendi/pyhandlexl
cd pyhandlexl
pip install -e .
```

## Development

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -e ".[dev]" --no-build-isolation
pytest
ruff check .
```

## License

MIT — see [LICENSE](LICENSE).
