# Add dependency injection for external services

## Goal
Decouple OpenAI and Outlook dependencies from UI for easier testing and replacement.

## Scope
- Introduce a simple DI layer or factory functions.
- Ensure services accept injected clients/adapters.

## Acceptance Criteria
- No direct OpenAI client creation in UI.
- Outlook adapter creation centralized in a factory.
- Tests can swap in stubs without monkeypatching globals.

## Tasks
- Create `src/web/services/providers.py`.
- Update intent and run services to use injected providers.
- Update tests to use providers.
