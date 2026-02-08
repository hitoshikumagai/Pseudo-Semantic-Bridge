# Refactor web_app.py for modularity and performance

## Status
Open

## Goal
Reduce `web_app.py` complexity and improve responsiveness by separating UI concerns, state orchestration, and heavy processing paths.

## Background
- `web_app.py` has grown into a large single file with mixed responsibilities.
- UI rendering, session state mutation, semantic payload sync, and action handlers are tightly coupled.
- This increases regression risk and makes performance tuning difficult.

## Scope
- Split `web_app.py` into focused modules under `src/web/` (tabs, hub views, state sync, helpers).
- Keep current user-facing behavior and labels unchanged.
- Reduce redundant recomputation and unnecessary rerenders in hub/input tabs.
- Add guardrails for widget-key-safe session state handling.

## Out of Scope
- Major UX redesign.
- Feature additions unrelated to refactoring/performance.

## Acceptance Criteria
- `web_app.py` is reduced to thin composition/bootstrapping logic.
- Tab rendering logic is moved to module-level functions with clear interfaces.
- Full regression test suite passes in conda env.
- No known Streamlit session-state mutation exceptions in current flows.
- Performance checks (load/render timings already tracked by app) show no regressions, ideally improved on hub-heavy paths.

## Tasks
- Define target module boundaries and migration order.
- Extract shared state and semantic sync helpers first.
- Extract each tab renderer incrementally (overview, rules, candidates, intent, hub, run).
- Add/adjust smoke tests for import, key actions, and widget-key safety.
- Run full regression: `conda run -n pseudo-semantic-bridge pytest -q`.
- Update docs with new module map and maintenance guidance.

## Notes
- Prefer small, reviewable commits per extraction phase to keep rollback easy.
