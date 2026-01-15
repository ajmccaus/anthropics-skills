# Claude Guidelines

## Before Push
Validate: `python -m py_compile *.py` and test imports.

## xlsx
- Excel COM first (Windows), LibreOffice fallback
- `data_only=True` = values, `data_only=False` = formulas
