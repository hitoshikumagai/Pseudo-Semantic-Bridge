# Fix diagram_mode session_state mutation after widget instantiation

## Status
Open

## Goal
Resolve Streamlit runtime failure in Semantic Layer Hub caused by mutating widget-bound session state after widget creation.

## Error
`streamlit.errors.StreamlitAPIException: st.session_state.diagram_mode cannot be modified after the widget with key diagram_mode is instantiated.`

## Scope
- Remove invalid direct assignment to `st.session_state["diagram_mode"]` after `st.radio(..., key="diagram_mode")`.
- Confirm Mermaid-related state handling follows the same Streamlit session-state rules.
- Keep current UX (table/mermaid toggle) unchanged.

## Reproduction
1. Open `web_app.py` in Streamlit.
2. Navigate to `Semantic Layer Hub`.
3. Render the `Visualization mode` radio (`key="diagram_mode"`).
4. App raises StreamlitAPIException due to post-instantiation mutation.

## Suspected Root Cause
- In `web_app.py`, the app sets:
- `diagram_mode = st.radio(..., key="diagram_mode")`
- `st.session_state["diagram_mode"] = diagram_mode`
- The second line mutates a widget-managed key after instantiation, which Streamlit forbids.

## Acceptance Criteria
- No StreamlitAPIException when opening or switching `Visualization mode`.
- `diagram_mode` persists correctly between reruns.
- Existing tests remain green.

## Tasks
- Refactor `diagram_mode` handling to rely on widget-managed state only.
- Add/adjust a smoke test covering hub visualization mode rendering.
- Run full regression in conda env: `conda run -n pseudo-semantic-bridge pytest -q`.
