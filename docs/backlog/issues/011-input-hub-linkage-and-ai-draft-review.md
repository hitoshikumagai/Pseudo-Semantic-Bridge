# Clarify input-to-hub linkage and AI draft review flow

## Status
Done (2026-02-08)

## Goal
Make each input tab explicitly show how user actions populate semantic hub fields, and add an AI-first draft rule workflow with human review.

## Scope
- Show clear input->hub mapping inside each semantic input tab.
- Add AI draft generation from user instruction and push drafts to review queue.
- Keep human approval as the final step before rules are merged.

## Acceptance Criteria
- Users can see where each tab writes inside `automation_assets`.
- Rule drafting starts from a human instruction and creates reviewable draft rules.
- Proposed rules remain review-first (select + merge) before becoming active rules.

## Tasks
- Add mapping panels in rules/candidates/intent tabs and semantic hub.
- Add "AI draft from instruction" workflow and append drafts to proposed rules.
- Document the new UX flow in semantic workflow docs.

## Resolution Notes
- Added explicit input-to-hub mapping panels in all semantic input tabs and hub tab.
- Added AI draft rule generation from human instruction with review queue insertion.
- Kept review-first merge flow (`Proposed Rules`) as the only path to active rules.
- Updated workflow docs to describe AI draft + human review loop.
- Added detailed rule/IR relationship and coverage tables in semantic hub.
- Added maintenance actions for unlinked proposed rules and missing IR-derived rules.
