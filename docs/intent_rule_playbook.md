# Intent Spec to Rule Playbook

## Purpose
Standardize the repeatable process for turning an intent specification into executable mail routing rules.

## When To Use
- New automation request needs a rule.
- Existing flow changes require updated routing.
- Ambiguous intent needs a proposal + review cycle.

## Inputs
- Intent Spec JSON (from UI or generated template/AI)
- Domain context (mailbox, subject filters, expected attachments)
- Quality threshold (min_quality_score)

## Steps
1. Create or update the Intent Spec.
2. Ensure `inputs.mail_rule` is present, or confirm steps are sufficient for inference.
3. Generate proposed rules from the Intent Spec.
4. Review proposed rules for correctness and safety.
5. Merge selected rules into the active ruleset.
6. Run tests and confirm pipeline behavior.

## UI Flow (Recommended)
- Intent Spec tab: generate spec and click `Generate Proposed Rule`.
- Rules tab: review Proposed Rules, select, and `Merge Selected`.
- Save Proposed and Save Rules as needed.

## Artifacts
- Schema: `specs/schema/intent_spec_v1.schema.json`
- Sample spec: `specs/accounting/invoice_intent_spec.sample.json`
- Proposed rules: `configs/accounting/mail_rules_proposed.json`
- Active rules: `configs/accounting/mail_business_rules.json`
- Implementation: `src/web/services/rules_service.py`

## Quality Gates
- `verification.min_quality_score` must meet or exceed the gate (default 0.8).
- If below gate, keep as proposed unless explicitly allowed.

## Conflict Rules
- Duplicate key conflicts are skipped.
- Subject filter conflicts are held for review.
- Invalid rule payloads are skipped.

## Validation
- Run `conda run -n pseudo-semantic-bridge pytest`.
- Confirm UI smoke tests pass.

## Checklist
- Intent Spec has `spec_id`, `steps`, `verification`, and `fallback`.
- `inputs.mail_rule` present or steps can be inferred.
- Proposed rule has `subject_filter` and `action_id`.
- Quality gate decision recorded (allow or reject).
- Rules saved and pipeline smoke is green.
