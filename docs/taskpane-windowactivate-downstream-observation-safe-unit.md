# TaskPane WindowActivate Downstream Observation Safe Unit

## Scope

This document records the Phase 4 R15 GO decision for WindowActivate downstream observation.

References:

- `docs/taskpane-display-recovery-freeze-line.md`
- `docs/taskpane-refresh-orchestration-target-boundary-map.md`
- `docs/taskpane-refresh-orchestration-responsibility-inventory.md`
- `docs/taskpane-display-recovery-current-state.md`

## GO Decision

R15 can be treated as a safe-first ownership separation when the separation is limited to downstream observation for WindowActivate-triggered refresh.

The separated boundary is `WindowActivateDownstreamObservation`.

It owns:

- `window-activate-display-refresh-trigger-start` emission
- `window-activate-display-refresh-trigger-outcome` emission
- WindowActivate display request trace fields used by downstream refresh observation
- the explicit observation facts that keep `Dispatched != completion`

`TaskPaneRefreshOrchestrationService` still owns:

- refresh path ordering
- precondition and normalized outcome order
- foreground guarantee handoff
- created CASE display session and completion check
- `case-display-completed` one-time emit

`WindowActivatePaneHandlingService` still owns:

- WindowActivate request creation
- case protection gate
- external workbook detection handoff
- case pane suppression gate
- dispatch to display / refresh entry

## Preserved Boundary

WindowActivate downstream observation remains an observation layer only.

- `WindowActivateDispatchOutcome.Dispatched` means the display request was handed off.
- `window-activate-display-refresh-trigger-start` means the downstream refresh path was entered from a WindowActivate request.
- `window-activate-display-refresh-trigger-outcome` means the downstream refresh path produced observable refresh facts.
- None of these traces mean display completion.
- `case-display-completed` remains the only created CASE display completion trace.

The trace source string remains `TaskPaneRefreshOrchestrationService` to avoid changing the existing trace contract while moving the ownership boundary.

## Freeze Line Confirmation

This safe unit does not change:

- `WindowActivate dispatch != completion`
- `case-display-completed` one-time emit
- display session boundary
- foreground outcome semantics
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
- WindowActivate dispatch gate ordering

## Test Contract

WindowActivate downstream observation tests freeze:

- downstream start trace keeps `displayCompletionOutcome=False`
- downstream outcome trace keeps `displayCompletionOutcome=False`
- downstream delegated recovery is reported as `activationAttempt=Delegated`, not completion
- non-WindowActivate display requests do not emit WindowActivate downstream observation
- display request trace fields keep the WindowActivate trigger role and non-owner facts
