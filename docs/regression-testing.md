# Regression Testing

## Scope
- Full repository regression is defined as `pytest -q` from project root.
- Run this in the project conda environment (`pseudo-semantic-bridge`) to ensure all dependencies are available.

## Recommended Commands
```bash
cd /Users/kumagaihitoshi/Documents/GitHub/Pseudo-Semantic-Bridge
conda run -n pseudo-semantic-bridge pytest -q
```

## Latest Run Record
- Date: 2026-02-08
- Command: `conda run -n pseudo-semantic-bridge pytest -q`
- Result: `129 passed in 3.17s`

## Troubleshooting
- Symptom: `ModuleNotFoundError: No module named 'pydantic'` during test collection.
- Cause: tests were executed outside the project conda environment.
- Action: use `conda run -n pseudo-semantic-bridge pytest -q` or activate that environment before running tests.
