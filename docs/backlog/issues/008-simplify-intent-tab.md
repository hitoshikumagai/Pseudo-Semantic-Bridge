# Simplify Intent Spec tab

## Goal
Rebuild the Intent Spec tab from a minimal, clear flow to reduce UI chaos.

## Scope
- Start with the smallest useful flow (conversation -> summary -> spec preview).
- Hide advanced inputs behind an "Advanced" toggle.
- Keep the Mail Rule mapping optional and separate.

## Acceptance Criteria
- The default view shows only the minimal inputs and a single primary action.
- Advanced fields are hidden by default.
- Users can still generate and preview an Intent Spec.
- No regression in existing data flow (spec preview + apply).

## Tasks
- Define minimal UI flow and remove redundant controls.
- Add an Advanced section for power users.
- Update tests for the new flow.
