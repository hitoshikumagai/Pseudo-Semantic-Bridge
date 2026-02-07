# Clarify service boundaries

## Goal
Ensure `run_service`, `rules_service`, `intent_service`, `quality_service` have clear responsibilities and no cross-import cycles.

## Scope
- Audit imports between services.
- Remove or refactor any circular dependencies.

## Acceptance Criteria
- No circular imports between service modules.
- Each service has a short module docstring describing responsibility.
- Tests pass.

## Tasks
- Identify shared utilities and extract to `src/web/services/common.py` if needed.
- Update tests if behavior changes.
