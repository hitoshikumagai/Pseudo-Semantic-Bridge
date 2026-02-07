# Reorganize web tests

## Goal
Group web-related tests under `tests/web/` to mirror UI/service structure.

## Scope
- Move `tests/test_web_app_logic.py`, `tests/test_web_app_smoke.py`, `tests/test_mail_workflow.py` into `tests/web/`.
- Update imports and any test discovery references.

## Acceptance Criteria
- Tests still run with `pytest -q`.
- Clear separation of web vs non-web tests.

## Tasks
- Create `tests/web/` package if needed.
- Update CI/test documentation.
