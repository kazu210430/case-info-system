# TaskPane Normalized Outcome Mapping Safe Unit

## Scope

This document records the Phase 4 R10/R11/R12 GO decision for normalized outcome mapping.

References:

- `docs/taskpane-display-recovery-freeze-line.md`
- `docs/taskpane-refresh-orchestration-target-boundary-map.md`
- `docs/taskpane-refresh-orchestration-responsibility-inventory.md`
- `docs/taskpane-display-recovery-current-state.md`

## GO Decision

R10/R11/R12 can be treated as a safe-first ownership separation when the separation is limited to normalized outcome mapping.

The separated boundary is `TaskPaneNormalizedOutcomeMapper`.

It owns:

- visibility recovery outcome mapping
- refresh source selection outcome mapping
- rebuild fallback outcome mapping
- normalized trace detail formatting for those outcomes
- refresh-source status to trace action mapping

`TaskPaneRefreshOrchestrationService` still owns:

- refresh / ready-show callback orchestration order
- trace emission
- attempt result aggregation
- foreground guarantee decision handoff
- created CASE display completion check

## Preserved Order

The normalized outcome order remains unchanged:

1. visibility recovery outcome
2. refresh source selection outcome
3. rebuild fallback outcome
4. foreground guarantee outcome
5. created CASE display completion check

This keeps R13 foreground decision and R14 `case-display-completed` ownership outside the R10/R11/R12 safe unit.

## Freeze Line Confirmation

This safe unit does not change:

- `pending != completion`
- `WindowActivate dispatch != completion`
- foreground outcome semantics
- `case-display-completed` one-time emit
- display session boundary
- trace names or trace meanings
- callback meaning
- retry sequencing

## Dangerous Adjacent Responsibilities Not Moved

The following responsibilities remain intentionally out of scope:

- R04/R14 display protocol session
- R05 callback/completion convergence
- R07/R08 pending fallback ownership
- R13 foreground decision
- `case-display-completed` emit owner

## Test Contract

Outcome normalization tests freeze:

- failed visibility recovery stays non-display-completable
- already-visible visibility recovery stays display-completable but skipped
- degraded visibility recovery facts stay visibility outcome facts
- refresh source not-reached remains not-reached
- rebuild fallback skipped remains a continuation outcome, not completion
- refresh-source trace action names remain unchanged
- visibility trace details keep completion source, attempt number, and display-completable facts
