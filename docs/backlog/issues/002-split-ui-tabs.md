# Split UI tabs into modules

## Goal
Move tab render logic out of `web_app.py` into `src/web/ui/tabs/*` modules.

## Scope
- Create one module per tab (Overview, Rules, Rule Builder, Intent Spec, Run).
- Keep `web_app.py` as a lightweight entrypoint that wires the tabs.

## Acceptance Criteria
- `web_app.py` is < 200 lines and only orchestrates tab calls.
- All tabs render identically to current behavior.
- Tests pass.

## Tasks
- Add tab modules under `src/web/ui/tabs/`.
- Introduce a `render_*` function for each tab.
- Update smoke tests if necessary.
