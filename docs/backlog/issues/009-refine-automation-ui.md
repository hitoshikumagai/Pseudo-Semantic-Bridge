# Refine two-tab automation UI

## Goal
Improve usability of the new two-tab layout (`Automation Summary`, `Automation Pipelines`) without changing core behavior.

## Scope
- Clarify information hierarchy in both tabs.
- Reduce cognitive load in `Automation Pipelines` by progressive disclosure.
- Keep all existing actions available (rules, candidates, intent spec, run).

## Acceptance Criteria
- The summary tab surfaces only high-value status first (KPIs, recent failures, active jobs).
- Pipeline sections have clear purpose, inputs, outputs, and primary action.
- Advanced controls are collapsed by default where possible.
- No regression in run/compile/rule-edit workflows.

## Tasks
- Define section order and default expand/collapse policy.
- Add short per-pipeline "What this does" and "When to use".
- Add visual separators and tighter labels for scanability.
- Update smoke tests to cover the restructured layout.
