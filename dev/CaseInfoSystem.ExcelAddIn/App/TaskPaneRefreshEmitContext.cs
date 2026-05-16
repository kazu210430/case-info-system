using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class TaskPaneRefreshEmitContextBuilder
    {
        private readonly Func<Excel.Workbook, string> _formatWorkbook;
        private readonly Func<Excel.Window, string> _formatWindow;

        internal TaskPaneRefreshEmitContextBuilder(
            Func<Excel.Workbook, string> formatWorkbook,
            Func<Excel.Window, string> formatWindow)
        {
            _formatWorkbook = formatWorkbook ?? throw new ArgumentNullException(nameof(formatWorkbook));
            _formatWindow = formatWindow ?? throw new ArgumentNullException(nameof(formatWindow));
        }

        internal TaskPaneRefreshEmitContext Build(TaskPaneRefreshEmitContextInput input)
        {
            if (input == null)
            {
                throw new ArgumentNullException(nameof(input));
            }

            return new TaskPaneRefreshEmitContext(
                input.Reason,
                input.SessionSnapshot,
                input.AttemptResult,
                input.CompletionSource,
                input.AttemptNumber,
                input.DisplayRequest,
                input.Workbook,
                input.Window,
                _formatWorkbook(input.Workbook),
                _formatWindow(input.Window));
        }
    }

    internal sealed class TaskPaneRefreshEmitContextInput
    {
        private TaskPaneRefreshEmitContextInput(
            string reason,
            CreatedCaseDisplaySessionSnapshot sessionSnapshot,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            TaskPaneDisplayRequest displayRequest,
            Excel.Workbook workbook,
            Excel.Window window)
        {
            Reason = reason;
            SessionSnapshot = sessionSnapshot;
            AttemptResult = attemptResult;
            CompletionSource = completionSource;
            AttemptNumber = attemptNumber;
            DisplayRequest = displayRequest;
            Workbook = workbook;
            Window = window;
        }

        internal string Reason { get; }

        internal CreatedCaseDisplaySessionSnapshot SessionSnapshot { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal string CompletionSource { get; }

        internal int? AttemptNumber { get; }

        internal TaskPaneDisplayRequest DisplayRequest { get; }

        internal Excel.Workbook Workbook { get; }

        internal Excel.Window Window { get; }

        internal static TaskPaneRefreshEmitContextInput ForCompletedSession(
            string reason,
            CreatedCaseDisplaySession session,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            TaskPaneDisplayRequest displayRequest,
            Excel.Workbook workbook,
            Excel.Window window)
        {
            return new TaskPaneRefreshEmitContextInput(
                reason,
                session == null ? null : session.ToSnapshot(),
                attemptResult,
                completionSource,
                attemptNumber,
                displayRequest,
                workbook,
                window);
        }

        internal static TaskPaneRefreshEmitContextInput ForCompletionResult(
            TaskPaneRefreshCompletionResult completionResult)
        {
            if (completionResult == null)
            {
                throw new ArgumentNullException(nameof(completionResult));
            }

            return new TaskPaneRefreshEmitContextInput(
                completionResult.Reason,
                completionResult.SessionSnapshot,
                completionResult.AttemptResult,
                completionResult.CompletionSource,
                completionResult.AttemptNumber,
                completionResult.DisplayRequest,
                completionResult.Workbook,
                completionResult.Window);
        }
    }

    internal sealed class TaskPaneRefreshEmitContext
    {
        internal TaskPaneRefreshEmitContext(
            string reason,
            CreatedCaseDisplaySessionSnapshot sessionSnapshot,
            TaskPaneRefreshAttemptResult attemptResult,
            string completionSource,
            int? attemptNumber,
            TaskPaneDisplayRequest displayRequest,
            Excel.Workbook workbook,
            Excel.Window window,
            string formattedWorkbook,
            string formattedWindow)
        {
            Reason = reason;
            SessionSnapshot = sessionSnapshot;
            SessionId = sessionSnapshot == null ? string.Empty : sessionSnapshot.SessionId;
            WorkbookFullName = sessionSnapshot == null ? string.Empty : sessionSnapshot.WorkbookFullName;
            AttemptResult = attemptResult;
            CompletionSource = completionSource;
            AttemptNumber = attemptNumber;
            DisplayRequest = displayRequest;
            Workbook = workbook;
            Window = window;
            FormattedWorkbook = formattedWorkbook ?? string.Empty;
            FormattedWindow = formattedWindow ?? string.Empty;
        }

        internal string Reason { get; }

        internal CreatedCaseDisplaySessionSnapshot SessionSnapshot { get; }

        internal string SessionId { get; }

        internal string WorkbookFullName { get; }

        internal TaskPaneRefreshAttemptResult AttemptResult { get; }

        internal string CompletionSource { get; }

        internal int? AttemptNumber { get; }

        internal TaskPaneDisplayRequest DisplayRequest { get; }

        internal Excel.Workbook Workbook { get; }

        internal Excel.Window Window { get; }

        internal string FormattedWorkbook { get; }

        internal string FormattedWindow { get; }
    }
}
