using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal static class TaskPaneNormalizedOutcomeMapper
    {
        internal static VisibilityRecoveryOutcome BuildVisibilityRecoveryOutcome(
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            WorkbookWindowVisibilityEnsureOutcome? ensureStatus = workbookWindowEnsureFacts == null
                ? (WorkbookWindowVisibilityEnsureOutcome?)null
                : workbookWindowEnsureFacts.Outcome;
            VisibilityRecoveryTargetKind targetKind = ResolveVisibilityRecoveryTargetKind(
                workbook,
                inputWindow,
                attemptResult);
            PaneVisibleSource paneVisibleSource = attemptResult.PaneVisibleSource;

            if (!attemptResult.IsRefreshSucceeded)
            {
                if (attemptResult.WasSkipped)
                {
                    return VisibilityRecoveryOutcome.Skipped(
                        "refreshSkipped",
                        isPaneVisible: false,
                        isDisplayCompletable: false,
                        targetKind: targetKind,
                        paneVisibleSource: paneVisibleSource,
                        workbookWindowEnsureStatus: ensureStatus,
                        fullRecoveryAttempted: attemptResult.PreContextRecoveryAttempted,
                        fullRecoverySucceeded: attemptResult.PreContextRecoverySucceeded);
                }

                return VisibilityRecoveryOutcome.Failed(
                    attemptResult.WasContextRejected ? "contextRejected" : "refreshFailed",
                    targetKind,
                    paneVisibleSource,
                    ensureStatus,
                    attemptResult.PreContextRecoveryAttempted,
                    attemptResult.PreContextRecoverySucceeded);
            }

            if (!attemptResult.IsPaneVisible)
            {
                return VisibilityRecoveryOutcome.Failed(
                    "paneVisible=false",
                    targetKind,
                    paneVisibleSource,
                    ensureStatus,
                    attemptResult.PreContextRecoveryAttempted,
                    attemptResult.PreContextRecoverySucceeded);
            }

            string degradedReason = ResolveVisibilityRecoveryDegradedReason(workbookWindowEnsureFacts, attemptResult);
            if (!string.IsNullOrWhiteSpace(degradedReason))
            {
                return VisibilityRecoveryOutcome.Degraded(
                    "paneVisibleWithDegradedRecoveryFacts",
                    targetKind,
                    paneVisibleSource,
                    ensureStatus,
                    attemptResult.PreContextRecoveryAttempted,
                    attemptResult.PreContextRecoverySucceeded,
                    degradedReason);
            }

            if (paneVisibleSource == PaneVisibleSource.AlreadyVisibleHost)
            {
                return VisibilityRecoveryOutcome.Skipped(
                    "alreadyVisible",
                    isPaneVisible: true,
                    isDisplayCompletable: true,
                    targetKind: VisibilityRecoveryTargetKind.AlreadyVisible,
                    paneVisibleSource: paneVisibleSource,
                    workbookWindowEnsureStatus: ensureStatus,
                    fullRecoveryAttempted: attemptResult.PreContextRecoveryAttempted,
                    fullRecoverySucceeded: attemptResult.PreContextRecoverySucceeded);
            }

            string completedReason = ensureStatus == WorkbookWindowVisibilityEnsureOutcome.MadeVisible
                ? "madeVisibleThenShown"
                : "paneVisible";
            if (attemptResult.IsRefreshCompleted)
            {
                completedReason = paneVisibleSource == PaneVisibleSource.ReusedShown
                    ? "reusedShown"
                    : "refreshedShown";
            }

            return VisibilityRecoveryOutcome.Completed(
                completedReason,
                targetKind,
                paneVisibleSource,
                ensureStatus,
                attemptResult.PreContextRecoveryAttempted,
                attemptResult.PreContextRecoverySucceeded);
        }

        internal static RefreshSourceSelectionOutcome BuildRefreshSourceSelectionOutcome(
            TaskPaneRefreshAttemptResult attemptResult)
        {
            return RefreshSourceSelectionOutcome.FromAttemptResult(attemptResult);
        }

        internal static RebuildFallbackOutcome BuildRebuildFallbackOutcome(
            TaskPaneRefreshAttemptResult attemptResult)
        {
            return RebuildFallbackOutcome.FromBuildResult(
                attemptResult == null ? null : attemptResult.SnapshotBuildResult);
        }

        internal static string FormatVisibilityRecoveryDetails(
            string reason,
            VisibilityRecoveryOutcome outcome,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts)
        {
            string details =
                "reason=" + (reason ?? string.Empty)
                + ",completionSource=" + (completionSource ?? string.Empty)
                + ",visibilityRecoveryStatus=" + outcome.Status.ToString()
                + ",visibilityRecoveryReason=" + outcome.Reason
                + ",visibilityRecoveryTerminal=" + outcome.IsTerminal.ToString()
                + ",visibilityRecoveryDisplayCompletable=" + outcome.IsDisplayCompletable.ToString()
                + ",visibilityRecoveryPaneVisible=" + outcome.IsPaneVisible.ToString()
                + ",visibilityRecoveryTargetKind=" + outcome.TargetKind.ToString()
                + ",visibilityPaneVisibleSource=" + outcome.PaneVisibleSource.ToString()
                + ",visibilityRecoveryDegradedReason=" + outcome.DegradedReason
                + ",refreshSucceeded=" + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ",refreshCompleted=" + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ",preContextFullRecoveryAttempted=" + outcome.FullRecoveryAttempted.ToString()
                + ",preContextFullRecoverySucceeded=" + FormatNullableBool(outcome.FullRecoverySucceeded);
            if (workbookWindowEnsureFacts != null)
            {
                details += ",workbookWindowEnsureStatus=" + workbookWindowEnsureFacts.Outcome.ToString()
                    + ",workbookWindowEnsureHwnd=" + workbookWindowEnsureFacts.WindowHwnd
                    + ",workbookWindowVisibleAfterSet=" + FormatNullableBool(workbookWindowEnsureFacts.VisibleAfterSet);
            }

            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return details;
        }

        internal static string FormatRefreshSourceSelectionAction(RefreshSourceSelectionOutcome outcome)
        {
            switch (outcome.Status)
            {
                case RefreshSourceSelectionOutcomeStatus.Selected:
                    return "refresh-source-selected";
                case RefreshSourceSelectionOutcomeStatus.DegradedSelected:
                    return "refresh-source-degraded";
                case RefreshSourceSelectionOutcomeStatus.FallbackSelected:
                    return "refresh-source-fallback";
                case RefreshSourceSelectionOutcomeStatus.RebuildRequired:
                    return "refresh-source-rebuild-required";
                case RefreshSourceSelectionOutcomeStatus.Failed:
                    return "refresh-source-failed";
                case RefreshSourceSelectionOutcomeStatus.NotReached:
                    return "refresh-source-not-reached";
                default:
                    return "refresh-source-unknown";
            }
        }

        internal static string FormatRefreshSourceSelectionDetails(
            string reason,
            RefreshSourceSelectionOutcome outcome,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber)
        {
            string details =
                "reason=" + (reason ?? string.Empty)
                + ",completionSource=" + (completionSource ?? string.Empty)
                + ",refreshSourceStatus=" + outcome.Status.ToString()
                + ",selectedSource=" + outcome.SelectedSource.ToString()
                + ",selectionReason=" + outcome.SelectionReason
                + ",fallbackReasons=" + outcome.FallbackReasons
                + ",refreshSourceTerminal=" + outcome.IsTerminal.ToString()
                + ",refreshSourceCanContinue=" + outcome.CanContinueRefresh.ToString()
                + ",cacheFallback=" + outcome.IsCacheFallback.ToString()
                + ",rebuildRequired=" + outcome.IsRebuildRequired.ToString()
                + ",masterListRebuildAttempted=" + outcome.MasterListRebuildAttempted.ToString()
                + ",masterListRebuildSucceeded=" + outcome.MasterListRebuildSucceeded.ToString()
                + ",snapshotTextAvailable=" + outcome.SnapshotTextAvailable.ToString()
                + ",updatedCaseSnapshotCache=" + outcome.UpdatedCaseSnapshotCache.ToString()
                + ",failureReason=" + outcome.FailureReason
                + ",degradedReason=" + outcome.DegradedReason
                + ",refreshSucceeded=" + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ",refreshCompleted=" + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ",paneVisible=" + (attemptResult != null && attemptResult.IsPaneVisible).ToString();
            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return details;
        }

        internal static string FormatRebuildFallbackDetails(
            string reason,
            RebuildFallbackOutcome outcome,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber)
        {
            string details =
                "reason=" + (reason ?? string.Empty)
                + ",completionSource=" + (completionSource ?? string.Empty)
                + ",rebuildFallbackStatus=" + outcome.Status.ToString()
                + ",rebuildFallbackRequired=" + outcome.IsRequired.ToString()
                + ",rebuildFallbackTerminal=" + outcome.IsTerminal.ToString()
                + ",rebuildFallbackCanContinue=" + outcome.CanContinueRefresh.ToString()
                + ",snapshotSource=" + outcome.SnapshotSource.ToString()
                + ",fallbackReasons=" + outcome.FallbackReasons
                + ",masterListRebuildAttempted=" + outcome.MasterListRebuildAttempted.ToString()
                + ",masterListRebuildSucceeded=" + outcome.MasterListRebuildSucceeded.ToString()
                + ",snapshotTextAvailable=" + outcome.SnapshotTextAvailable.ToString()
                + ",updatedCaseSnapshotCache=" + outcome.UpdatedCaseSnapshotCache.ToString()
                + ",failureReason=" + outcome.FailureReason
                + ",degradedReason=" + outcome.DegradedReason
                + ",outcomeReason=" + outcome.Reason
                + ",refreshSucceeded=" + (attemptResult != null && attemptResult.IsRefreshSucceeded).ToString()
                + ",refreshCompleted=" + (attemptResult != null && attemptResult.IsRefreshCompleted).ToString()
                + ",paneVisible=" + (attemptResult != null && attemptResult.IsPaneVisible).ToString();
            if (attemptNumber.HasValue)
            {
                details += ",attempt=" + attemptNumber.Value.ToString(CultureInfo.InvariantCulture);
            }

            return details;
        }

        private static VisibilityRecoveryTargetKind ResolveVisibilityRecoveryTargetKind(
            Excel.Workbook workbook,
            Excel.Window inputWindow,
            TaskPaneRefreshAttemptResult attemptResult)
        {
            if (attemptResult != null && attemptResult.PaneVisibleSource == PaneVisibleSource.AlreadyVisibleHost)
            {
                return VisibilityRecoveryTargetKind.AlreadyVisible;
            }

            if (workbook == null
                && inputWindow == null
                && attemptResult != null
                && attemptResult.ForegroundWorkbook == null
                && attemptResult.ForegroundWindow == null)
            {
                return attemptResult.ForegroundContext != null
                    || attemptResult.IsRefreshSucceeded
                    || attemptResult.PreContextRecoveryAttempted
                    ? VisibilityRecoveryTargetKind.ActiveWorkbookFallback
                    : VisibilityRecoveryTargetKind.NoKnownTarget;
            }

            if (workbook != null
                || inputWindow != null
                || (attemptResult != null && (attemptResult.ForegroundWorkbook != null || attemptResult.ForegroundWindow != null)))
            {
                return VisibilityRecoveryTargetKind.ExplicitWorkbookWindow;
            }

            return VisibilityRecoveryTargetKind.NoKnownTarget;
        }

        private static string ResolveVisibilityRecoveryDegradedReason(
            WorkbookWindowVisibilityEnsureFacts workbookWindowEnsureFacts,
            TaskPaneRefreshAttemptResult attemptResult)
        {
            if (workbookWindowEnsureFacts != null)
            {
                switch (workbookWindowEnsureFacts.Outcome)
                {
                    case WorkbookWindowVisibilityEnsureOutcome.WorkbookMissing:
                    case WorkbookWindowVisibilityEnsureOutcome.WindowUnresolved:
                    case WorkbookWindowVisibilityEnsureOutcome.VisibilityReadFailed:
                    case WorkbookWindowVisibilityEnsureOutcome.Failed:
                        return "workbookWindowEnsure=" + workbookWindowEnsureFacts.Outcome.ToString();
                    case WorkbookWindowVisibilityEnsureOutcome.MadeVisible:
                        if (workbookWindowEnsureFacts.VisibleAfterSet != true)
                        {
                            return "workbookWindowEnsureVisibleAfterSet="
                                + (workbookWindowEnsureFacts.VisibleAfterSet.HasValue
                                    ? workbookWindowEnsureFacts.VisibleAfterSet.Value.ToString()
                                    : "null");
                        }

                        break;
                }
            }

            if (attemptResult != null
                && attemptResult.PreContextRecoveryAttempted
                && attemptResult.PreContextRecoverySucceeded == false)
            {
                return "fullRecoveryReturnedFalse";
            }

            return string.Empty;
        }

        private static string FormatNullableBool(bool? value)
        {
            return value.HasValue ? value.Value.ToString() : string.Empty;
        }
    }
}
