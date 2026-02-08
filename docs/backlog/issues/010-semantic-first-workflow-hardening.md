# Harden semantic-first workflow

## Status
Done (2026-02-08)

## Goal
Complete the shift to a semantic-first app flow where all design tabs are inputs and `Run` is execution-only.

## Scope
- Remove remaining ambiguous behavior between tab-local state and semantic aggregated state.
- Add explicit run gating based on semantic readiness and required assets.
- Improve visibility of what was projected from semantic assets at runtime.

## Acceptance Criteria
- `Run` is blocked with clear messages when required semantic assets are missing.
- Users can see which assets are currently sourced from `automation_assets`.
- No regression in existing smoke behavior.

## Tasks
- Add readiness gates for `Run` (intent spec + rules + semantic minimum fields).
- Add a compact "semantic source of truth" status panel to Overview.
- Add tests for projection path: semantic -> tab state -> runtime preparation.
- Document the new run prerequisites in workflow docs.

## Resolution Notes
- Implemented run gating with explicit prerequisite messages in `Run Automation`.
- Added a compact "Semantic Source Of Truth" panel in `Semantic Overview`.
- Added smoke coverage for blocked run and semantic projection path.
- Updated semantic workflow docs with run prerequisite rules.
