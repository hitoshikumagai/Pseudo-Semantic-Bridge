# Centralize session state

## Goal
Collect `st.session_state` initialization and defaults into a single state module.

## Scope
- Introduce `src/web/ui/state.py` (or similar) with init helpers.
- Replace scattered session_state initialization in UI.

## Acceptance Criteria
- All session keys are initialized in one place.
- UI tabs read from state helpers rather than manual defaults.
- Tests pass.

## Tasks
- Create `init_state()` function.
- Update tab modules to call `init_state()` once.
- Add state unit tests if needed.
