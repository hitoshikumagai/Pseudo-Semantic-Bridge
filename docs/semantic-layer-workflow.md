# Semantic Layer First Workflow

## Intent
The app now treats the semantic layer definition as the primary source of truth.
Each design tab contributes inputs into one aggregated semantic payload, and the final run uses that payload.

## UI Flow
1. `Rules` updates rule assets in `automation_assets.rules` and `automation_assets.proposed_rules`.
2. `Design: Rule Builder` updates `automation_assets.candidate_meta` and `automation_assets.candidate_rows`.
3. `Design: Intent Spec` updates `automation_assets.intent_spec` and `automation_assets.intent_spec_source`.
4. `Design: Semantic Layer` edits business/metadata/governance sections and previews the full aggregated payload.
5. `Run` prepares runtime inputs from the semantic payload and then executes compile/run.

## Data Model
Semantic layer file:
- `configs/accounting/semantic_layer_definition.json`

Important sections:
- `purpose`
- `technical_metadata`
- `business_semantics`
- `federation`
- `active_metadata`
- `ownership`
- `automation_assets`

`automation_assets` is the bridge between design tabs and run execution.

## Runtime Behavior
- On app load, semantic assets are projected into working tab state when available.
- During interaction, working state is continuously re-aggregated back into `semantic_layer_spec` in session.
- `Run` calls runtime preparation, writes effective rules to `RULES_PATH`, and then executes the pipeline.

## Run Prerequisites
`Run Automation` is gated until the following semantic prerequisites are satisfied:
- `purpose.objective_statement` is non-empty.
- `purpose.priority_domain` is non-empty.
- `automation_assets.rules` has at least one rule.
- `automation_assets.intent_spec.spec_id` exists.

If any prerequisite is missing, the UI shows blocking messages and does not start compile/run.
