# Remove app_logic facade

## Goal
Stop using `src/web/app_logic.py` as a facade and import services directly from `src/web/services/*`.

## Scope
- Update `web_app.py` and any other modules to import services directly.
- Remove facade re-exports and keep only what is needed for backward compatibility (if any).

## Acceptance Criteria
- No runtime imports from `src/web/app_logic.py` in production code.
- All tests pass.

## Tasks
- Replace imports in `web_app.py` to target service modules.
- Remove or slim down `src/web/app_logic.py`.
- Update tests referencing `app_logic` if needed.
