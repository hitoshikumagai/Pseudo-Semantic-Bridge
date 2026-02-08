# Add table/diagram visualization and AI help for semantic maintenance

## Status
Done (2026-02-08)

## Goal
Improve semantic hub usability for existing assets by adding system-design style visualization and AI-assisted editing.

## Scope
- Add table view and Mermaid view for input-hub-run relationships.
- Add AI-assisted Mermaid generation/editing from natural language prompts.
- Add AI help actions to intent-related input screens.

## Acceptance Criteria
- Hub can switch between table and Mermaid representations.
- Users can generate Mermaid flow from AI prompt and keep editing manually.
- Intent-related tabs expose at least one AI help action each.
- Existing smoke tests remain green.

## Tasks
- Build relationship tables for nodes/edges and show in hub.
- Build Mermaid generator from semantic payload + AI prompt.
- Add AI help buttons for candidate instruction and intent workflow.
- Update docs and tests.

## Resolution Notes
- Added table and Mermaid visualization modes in semantic hub.
- Added AI Mermaid draft generation from natural language prompt.
- Added AI help buttons in candidate and intent tabs.
- Updated tests and workflow docs.
