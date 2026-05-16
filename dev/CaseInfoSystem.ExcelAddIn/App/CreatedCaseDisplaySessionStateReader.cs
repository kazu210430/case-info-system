using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class CreatedCaseDisplaySessionStateReader
    {
        internal CreatedCaseDisplaySessionStartDecision DecideStart(CreatedCaseDisplaySessionStartInput input)
        {
            if (input == null || !input.IsCreatedCaseDisplayReason)
            {
                return CreatedCaseDisplaySessionStartDecision.Blocked("reasonNotCreatedCaseDisplay");
            }

            if (string.IsNullOrWhiteSpace(input.WorkbookFullName))
            {
                return CreatedCaseDisplaySessionStartDecision.Blocked("workbookFullName=nullOrEmpty");
            }

            return CreatedCaseDisplaySessionStartDecision.Allowed();
        }

        internal CreatedCaseDisplaySessionSnapshot ResolveForCompletion(CreatedCaseDisplaySessionResolutionInput input)
        {
            if (input == null || !input.IsCreatedCaseDisplayReason)
            {
                return null;
            }

            CreatedCaseDisplaySessionSnapshot singleSession = null;
            int activeSessionCount = 0;
            IEnumerable<CreatedCaseDisplaySessionSnapshot> activeSessions = input.ActiveSessions;
            if (activeSessions == null)
            {
                return null;
            }

            foreach (CreatedCaseDisplaySessionSnapshot activeSession in activeSessions)
            {
                if (activeSession == null)
                {
                    continue;
                }

                activeSessionCount++;
                singleSession = activeSession;
                if (!string.IsNullOrWhiteSpace(input.WorkbookFullName)
                    && string.Equals(
                        activeSession.WorkbookFullName,
                        input.WorkbookFullName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return activeSession;
                }
            }

            return activeSessionCount == 1 ? singleSession : null;
        }
    }

    internal sealed class CreatedCaseDisplaySessionStartInput
    {
        internal CreatedCaseDisplaySessionStartInput(bool isCreatedCaseDisplayReason, string workbookFullName)
        {
            IsCreatedCaseDisplayReason = isCreatedCaseDisplayReason;
            WorkbookFullName = workbookFullName ?? string.Empty;
        }

        internal bool IsCreatedCaseDisplayReason { get; }

        internal string WorkbookFullName { get; }
    }

    internal struct CreatedCaseDisplaySessionStartDecision
    {
        private CreatedCaseDisplaySessionStartDecision(bool shouldStart, string blockedReason)
        {
            ShouldStart = shouldStart;
            BlockedReason = blockedReason ?? string.Empty;
        }

        internal bool ShouldStart { get; }

        internal string BlockedReason { get; }

        internal static CreatedCaseDisplaySessionStartDecision Allowed()
        {
            return new CreatedCaseDisplaySessionStartDecision(true, string.Empty);
        }

        internal static CreatedCaseDisplaySessionStartDecision Blocked(string blockedReason)
        {
            return new CreatedCaseDisplaySessionStartDecision(false, blockedReason);
        }
    }

    internal sealed class CreatedCaseDisplaySessionResolutionInput
    {
        internal CreatedCaseDisplaySessionResolutionInput(
            bool isCreatedCaseDisplayReason,
            string workbookFullName,
            IEnumerable<CreatedCaseDisplaySessionSnapshot> activeSessions)
        {
            IsCreatedCaseDisplayReason = isCreatedCaseDisplayReason;
            WorkbookFullName = workbookFullName ?? string.Empty;
            ActiveSessions = activeSessions;
        }

        internal bool IsCreatedCaseDisplayReason { get; }

        internal string WorkbookFullName { get; }

        internal IEnumerable<CreatedCaseDisplaySessionSnapshot> ActiveSessions { get; }
    }

    internal sealed class CreatedCaseDisplaySessionSnapshot
    {
        internal CreatedCaseDisplaySessionSnapshot(
            string sessionId,
            string workbookFullName,
            string reason,
            bool isCompleted)
        {
            SessionId = sessionId ?? string.Empty;
            WorkbookFullName = workbookFullName ?? string.Empty;
            Reason = reason ?? string.Empty;
            IsCompleted = isCompleted;
        }

        internal string SessionId { get; }

        internal string WorkbookFullName { get; }

        internal string Reason { get; }

        internal bool IsCompleted { get; }
    }

    internal sealed class CreatedCaseDisplaySession
    {
        internal CreatedCaseDisplaySession(string sessionId, string workbookFullName, string reason)
        {
            SessionId = sessionId ?? string.Empty;
            WorkbookFullName = workbookFullName ?? string.Empty;
            Reason = reason ?? string.Empty;
        }

        internal string SessionId { get; }

        internal string WorkbookFullName { get; }

        internal string Reason { get; }

        internal bool IsCompleted { get; set; }

        internal CreatedCaseDisplaySessionSnapshot ToSnapshot()
        {
            return new CreatedCaseDisplaySessionSnapshot(
                SessionId,
                WorkbookFullName,
                Reason,
                IsCompleted);
        }
    }
}
