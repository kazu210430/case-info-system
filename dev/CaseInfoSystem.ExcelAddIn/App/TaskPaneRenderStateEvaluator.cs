using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneDisplayEntryState
    {
        internal TaskPaneDisplayEntryState(
            bool hasTargetWindow,
            bool hasResolvableWindowKey,
            bool hasManagedPane,
            bool hasExistingHost,
            bool isSameWorkbook,
            bool isRenderSignatureCurrent)
        {
            HasTargetWindow = hasTargetWindow;
            HasResolvableWindowKey = hasResolvableWindowKey;
            HasManagedPane = hasManagedPane;
            HasExistingHost = hasExistingHost;
            IsSameWorkbook = isSameWorkbook;
            IsRenderSignatureCurrent = isRenderSignatureCurrent;
        }

        internal bool HasTargetWindow { get; }

        internal bool HasResolvableWindowKey { get; }

        internal bool HasManagedPane { get; }

        internal bool HasExistingHost { get; }

        internal bool IsSameWorkbook { get; }

        internal bool IsRenderSignatureCurrent { get; }
    }

    internal sealed class TaskPaneRenderStateEvaluation
    {
        internal TaskPaneRenderStateEvaluation(string renderSignature, bool isRenderRequired)
        {
            RenderSignature = renderSignature ?? string.Empty;
            IsRenderRequired = isRenderRequired;
        }

        internal string RenderSignature { get; }

        internal bool IsRenderRequired { get; }
    }

    internal static class TaskPaneRenderStateEvaluator
    {
        internal static TaskPaneDisplayEntryState EvaluateDisplayEntryState(
            ExcelInteropService excelInteropService,
            IDictionary<string, TaskPaneHost> hostsByWindowKey,
            Excel.Workbook workbook,
            Excel.Window window)
        {
            if (excelInteropService == null)
            {
                throw new ArgumentNullException(nameof(excelInteropService));
            }

            if (hostsByWindowKey == null)
            {
                throw new ArgumentNullException(nameof(hostsByWindowKey));
            }

            bool hasTargetWindow = window != null;
            if (!hasTargetWindow)
            {
                return new TaskPaneDisplayEntryState(
                    hasTargetWindow: false,
                    hasResolvableWindowKey: false,
                    hasManagedPane: false,
                    hasExistingHost: false,
                    isSameWorkbook: false,
                    isRenderSignatureCurrent: false);
            }

            string windowKey = SafeGetWindowKey(window);
            bool hasResolvableWindowKey = !string.IsNullOrWhiteSpace(windowKey);
            if (!hasResolvableWindowKey
                || !hostsByWindowKey.TryGetValue(windowKey, out TaskPaneHost host))
            {
                return new TaskPaneDisplayEntryState(
                    hasTargetWindow: true,
                    hasResolvableWindowKey: hasResolvableWindowKey,
                    hasManagedPane: false,
                    hasExistingHost: false,
                    isSameWorkbook: false,
                    isRenderSignatureCurrent: false);
            }

            string workbookFullName = workbook == null ? string.Empty : excelInteropService.GetWorkbookFullName(workbook);
            bool isSameWorkbook =
                !string.IsNullOrWhiteSpace(workbookFullName)
                && string.Equals(host.WorkbookFullName, workbookFullName, StringComparison.OrdinalIgnoreCase);
            if (!isSameWorkbook)
            {
                return new TaskPaneDisplayEntryState(
                    hasTargetWindow: true,
                    hasResolvableWindowKey: true,
                    hasManagedPane: true,
                    hasExistingHost: true,
                    isSameWorkbook: false,
                    isRenderSignatureCurrent: false);
            }

            WorkbookRole role = GetHostedWorkbookRole(host);
            if (role == WorkbookRole.Unknown)
            {
                return new TaskPaneDisplayEntryState(
                    hasTargetWindow: true,
                    hasResolvableWindowKey: true,
                    hasManagedPane: true,
                    hasExistingHost: true,
                    isSameWorkbook: true,
                    isRenderSignatureCurrent: false);
            }

            string renderSignature = BuildRenderSignature(
                excelInteropService,
                new WorkbookContext(
                    workbook,
                    window,
                    role,
                    excelInteropService.TryGetDocumentProperty(workbook, "SYSTEM_ROOT"),
                    workbookFullName,
                    excelInteropService.GetActiveSheetCodeName(workbook)));
            bool isRenderSignatureCurrent =
                !string.IsNullOrWhiteSpace(host.LastRenderSignature)
                && string.Equals(host.LastRenderSignature, renderSignature, StringComparison.Ordinal);
            return new TaskPaneDisplayEntryState(
                hasTargetWindow: true,
                hasResolvableWindowKey: true,
                hasManagedPane: true,
                hasExistingHost: true,
                isSameWorkbook: true,
                isRenderSignatureCurrent: isRenderSignatureCurrent);
        }

        internal static TaskPaneRenderStateEvaluation EvaluateRenderState(
            ExcelInteropService excelInteropService,
            TaskPaneHost host,
            WorkbookContext context)
        {
            string renderSignature = BuildRenderSignature(excelInteropService, context);
            bool isRenderRequired = host == null
                || !string.Equals(host.LastRenderSignature, renderSignature, StringComparison.Ordinal);
            return new TaskPaneRenderStateEvaluation(renderSignature, isRenderRequired);
        }

        internal static string BuildRenderSignature(ExcelInteropService excelInteropService, WorkbookContext context)
        {
            if (excelInteropService == null)
            {
                throw new ArgumentNullException(nameof(excelInteropService));
            }

            if (context == null)
            {
                return string.Empty;
            }

            string caseListRegistered = string.Empty;
            string snapshotCacheCount = string.Empty;
            if (context.Role == WorkbookRole.Case && context.Workbook != null)
            {
                caseListRegistered = excelInteropService.TryGetDocumentProperty(context.Workbook, "CASELIST_REGISTERED") ?? string.Empty;
                snapshotCacheCount = excelInteropService.TryGetDocumentProperty(context.Workbook, "TASKPANE_SNAPSHOT_CACHE_COUNT") ?? string.Empty;
            }

            return string.Join(
                "|",
                context.Role.ToString(),
                context.WorkbookFullName ?? string.Empty,
                context.ActiveSheetCodeName ?? string.Empty,
                caseListRegistered,
                snapshotCacheCount);
        }

        private static WorkbookRole GetHostedWorkbookRole(TaskPaneHost host)
        {
            if (host == null || host.Control == null)
            {
                return WorkbookRole.Unknown;
            }

            if (host.Control is DocumentButtonsControl)
            {
                return WorkbookRole.Case;
            }

            if (host.Control is KernelNavigationControl)
            {
                return WorkbookRole.Kernel;
            }

            if (host.Control is AccountingNavigationControl)
            {
                return WorkbookRole.Accounting;
            }

            return WorkbookRole.Unknown;
        }

        private static string SafeGetWindowKey(Excel.Window window)
        {
            try
            {
                return window == null ? string.Empty : Convert.ToString(window.Hwnd) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
