# ready-show / recovery observation points (2026-05-08)

## Purpose

This note records the additional observation points introduced for the third safe unit.
It is observation-only work for the unstable display sequence after template update:

- new CASE creation
- first visible display
- reopen after first display

This note does not declare a root cause.
It does not change UI behavior, recovery conditions, retry counts, visibility control, or ready-show timing.

Related source-of-truth docs:

- `docs/flows.md`
- `docs/taskpane-refresh-policy.md`
- `docs/codex-review-first-safe-unit-plan.md`

## Correlation fields

`NewCaseVisibilityObservation.FormatCorrelationFields(...)` is the common correlation entrypoint for the added logs.

Added correlation fields:

- `observationSessionId`
  - Present while `NewCaseVisibilityObservation` is tracking the new CASE flow.
  - Format: `NCO-xxxx`
- `traceKey`
  - Stable flow key for the workbook.
  - Same value as `workbookId`
- `workbookId`
  - Stable ID generated as `WB-XXXXXXXXXXXX`
  - Source: uppercase workbook path
  - Fallback source: workbook name
- `workbookPath`
- `workbookName`
- `systemRoot`
- `taskPaneMasterVersion`
- `taskPaneBaseMasterVersion`

## Added observation points

### CASE creation to workbook open

- `KernelCasePresentationService.OpenCreatedCase`
  - `case-workbook-open-started`
  - `case-workbook-open-completed`
  - Plain log: `Created CASE workbook open started...`

### Workbook event order

- `WorkbookLifecycleCoordinator.OnWorkbookOpen`
  - `WorkbookOpen-event`
  - Plain log now includes correlation fields
- `WorkbookLifecycleCoordinator.OnWorkbookActivate`
  - `WorkbookActivate-event`
  - Plain log now includes correlation fields
- `ThisAddIn.HandleWindowActivateEvent`
  - `WindowActivate-event`
  - Plain log now includes correlation fields

### ready-show flow

- `TaskPaneRefreshOrchestrationService`
  - `ready-show-fallback-handoff`
  - `ready-show-enqueued`
- `WorkbookTaskPaneReadyShowAttemptWorker`
  - `ready-show-attempt`
  - `ready-show-attempt-result`
  - `ready-show-attempts-exhausted`

### refresh / foreground recovery flow

- `TaskPaneRefreshCoordinator`
  - `taskpane-refresh-started`
  - `taskpane-refresh-completed`
  - `foreground-recovery-decision`
  - `final-foreground-guarantee-started`
  - `final-foreground-guarantee-completed`
  - The refresh logs also carry `refreshSource=<reason>`

### window visibility / window recovery flow

- `WorkbookWindowVisibilityService`
  - Plain log: `Workbook window visibility recovery evaluated...`
  - Outcome includes whether visibility was changed and the elapsed time
- `ExcelWindowRecoveryService`
  - Plain log: `Excel window recovery evaluated...`
  - Plain log: `Excel window recovery mutation trace...`
  - Both logs now carry the same correlation fields as the CASE-side logs

### snapshot source / rebuild flow

- `TaskPaneSnapshotBuilderService`
  - `source=CaseCache`
  - `source=BaseCache`
  - `source=BaseCacheFallback`
  - `source=MasterListRebuild`
  - `Task pane snapshot rebuild fallback selected...`
  - `Task pane snapshot MasterListRebuild started...`

The snapshot logs now keep the following together:

- source
- `caseMasterVersion`
- `embeddedMasterVersion`
- `latestMasterVersion`
- `rebuildFallback`
- `resolvedMasterPath`
- common correlation fields

## Expected manual trace order

For a single new CASE flow, the main order to inspect is:

1. `case-workbook-open-started`
2. `case-workbook-open-completed`
3. `WorkbookOpen-event`
4. `WorkbookActivate-event`
5. `WindowActivate-event`
6. `ready-show-enqueued`
7. `ready-show-attempt`
8. `taskpane-refresh-started`
9. `taskpane-refresh-completed`
10. `foreground-recovery-decision`
11. `final-foreground-guarantee-started`
12. `final-foreground-guarantee-completed`

In parallel, inspect whether the same `traceKey` shows:

- `Workbook window visibility recovery evaluated...`
- `Excel window recovery evaluated...`
- `Excel window recovery mutation trace...`
- `source=CaseCache|BaseCache|BaseCacheFallback|MasterListRebuild`

## What to check next in runtime

- Whether `WorkbookOpen -> WorkbookActivate -> WindowActivate` always occurs in the same order for the problematic flow
- Whether `ready-show-enqueued` happens before or after the first visibility recovery
- Whether `taskpane-refresh-started` is reached from the first ready-show attempt or only from a later retry
- Whether `foreground-recovery-decision` starts after a successful refresh or is skipped
- Whether reopen switches snapshot source from `CaseCache` to `BaseCache` or `MasterListRebuild`
- Whether `taskPaneMasterVersion` and `taskPaneBaseMasterVersion` diverge when display instability appears
- Whether the same `traceKey` shows different workbook event order between initial display and reopen
