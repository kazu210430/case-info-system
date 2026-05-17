using System;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal sealed class CaseWorkbookHiddenAppCacheSlotDecisionService
    {
        internal CaseWorkbookHiddenAppCacheAcquisitionDecision DecideAcquisition(
            bool slotPresent,
            bool slotInUse,
            CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts)
        {
            if (!slotPresent)
            {
                return CaseWorkbookHiddenAppCacheAcquisitionDecision.CreateNew("cache-empty");
            }

            if (slotInUse)
            {
                return CaseWorkbookHiddenAppCacheAcquisitionDecision.BypassInUse("hiddenApplicationCacheInUse");
            }

            CaseWorkbookHiddenAppLifecycleFacts safeFacts = lifecycleFacts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
            if (!safeFacts.IsReusable)
            {
                return CaseWorkbookHiddenAppCacheAcquisitionDecision.DisposeUnhealthyAndCreate(
                    "acquire-health-check-failed",
                    "cache-replaced-after-unhealthy");
            }

            return CaseWorkbookHiddenAppCacheAcquisitionDecision.Reuse("cache-reusable");
        }

        internal CaseWorkbookHiddenAppCacheReturnSlotDecision DecideReturnSlot(
            bool slotPresent,
            bool sameApplication,
            bool isOwnedByCache)
        {
            return slotPresent && sameApplication && isOwnedByCache
                ? CaseWorkbookHiddenAppCacheReturnSlotDecision.Returnable()
                : CaseWorkbookHiddenAppCacheReturnSlotDecision.NotCacheSlot("cache-slot-mismatch");
        }

        internal CaseWorkbookHiddenAppCacheReturnHealthDecision DecideReturnHealth(
            CaseWorkbookHiddenAppLifecycleFacts lifecycleFacts)
        {
            CaseWorkbookHiddenAppLifecycleFacts safeFacts = lifecycleFacts ?? CaseWorkbookHiddenAppLifecycleFacts.Empty;
            return safeFacts.IsApplicationStateHealthy
                ? CaseWorkbookHiddenAppCacheReturnHealthDecision.Healthy()
                : CaseWorkbookHiddenAppCacheReturnHealthDecision.Unhealthy("return-to-idle-health-check-failed");
        }

        internal CaseWorkbookHiddenAppCacheDisposalDecision DecideDisposal(
            bool slotPresent,
            bool isOwnedByCache,
            bool applicationPresent)
        {
            if (!slotPresent)
            {
                return CaseWorkbookHiddenAppCacheDisposalDecision.NotRequired();
            }

            if (!isOwnedByCache)
            {
                return CaseWorkbookHiddenAppCacheDisposalDecision.SkipNotOwned(applicationPresent);
            }

            return CaseWorkbookHiddenAppCacheDisposalDecision.DisposeOwned(applicationPresent);
        }
    }

    internal sealed class CaseWorkbookHiddenAppCacheAcquisitionDecision
    {
        private CaseWorkbookHiddenAppCacheAcquisitionDecision(
            string action,
            string reason,
            string nextAcquisitionReason)
        {
            Action = action ?? string.Empty;
            Reason = reason ?? string.Empty;
            NextAcquisitionReason = nextAcquisitionReason ?? string.Empty;
        }

        internal string Action { get; }

        internal string Reason { get; }

        internal string NextAcquisitionReason { get; }

        internal bool ShouldBypassCache
        {
            get { return string.Equals(Action, "bypass-in-use", StringComparison.Ordinal); }
        }

        internal bool ShouldDisposeSlot
        {
            get { return string.Equals(Action, "dispose-unhealthy-and-create", StringComparison.Ordinal); }
        }

        internal bool ShouldReuseSlot
        {
            get { return string.Equals(Action, "reuse", StringComparison.Ordinal); }
        }

        internal static CaseWorkbookHiddenAppCacheAcquisitionDecision CreateNew(string reason)
        {
            return new CaseWorkbookHiddenAppCacheAcquisitionDecision("create-new", reason, reason);
        }

        internal static CaseWorkbookHiddenAppCacheAcquisitionDecision BypassInUse(string reason)
        {
            return new CaseWorkbookHiddenAppCacheAcquisitionDecision("bypass-in-use", reason, reason);
        }

        internal static CaseWorkbookHiddenAppCacheAcquisitionDecision DisposeUnhealthyAndCreate(
            string reason,
            string nextAcquisitionReason)
        {
            return new CaseWorkbookHiddenAppCacheAcquisitionDecision(
                "dispose-unhealthy-and-create",
                reason,
                nextAcquisitionReason);
        }

        internal static CaseWorkbookHiddenAppCacheAcquisitionDecision Reuse(string reason)
        {
            return new CaseWorkbookHiddenAppCacheAcquisitionDecision("reuse", reason, reason);
        }
    }

    internal sealed class CaseWorkbookHiddenAppCacheReturnSlotDecision
    {
        private CaseWorkbookHiddenAppCacheReturnSlotDecision(bool canReturn, string reason)
        {
            CanReturn = canReturn;
            Reason = reason ?? string.Empty;
        }

        internal bool CanReturn { get; }

        internal string Reason { get; }

        internal static CaseWorkbookHiddenAppCacheReturnSlotDecision Returnable()
        {
            return new CaseWorkbookHiddenAppCacheReturnSlotDecision(true, string.Empty);
        }

        internal static CaseWorkbookHiddenAppCacheReturnSlotDecision NotCacheSlot(string reason)
        {
            return new CaseWorkbookHiddenAppCacheReturnSlotDecision(false, reason);
        }
    }

    internal sealed class CaseWorkbookHiddenAppCacheReturnHealthDecision
    {
        private CaseWorkbookHiddenAppCacheReturnHealthDecision(bool canReturn, string reason)
        {
            CanReturn = canReturn;
            Reason = reason ?? string.Empty;
        }

        internal bool CanReturn { get; }

        internal string Reason { get; }

        internal static CaseWorkbookHiddenAppCacheReturnHealthDecision Healthy()
        {
            return new CaseWorkbookHiddenAppCacheReturnHealthDecision(true, string.Empty);
        }

        internal static CaseWorkbookHiddenAppCacheReturnHealthDecision Unhealthy(string reason)
        {
            return new CaseWorkbookHiddenAppCacheReturnHealthDecision(false, reason);
        }
    }

    internal sealed class CaseWorkbookHiddenAppCacheDisposalDecision
    {
        private CaseWorkbookHiddenAppCacheDisposalDecision(
            bool disposalRequired,
            bool shouldQuitApplication,
            bool shouldReleaseApplication,
            bool retainedInstancePresent,
            string reason)
        {
            DisposalRequired = disposalRequired;
            ShouldQuitApplication = shouldQuitApplication;
            ShouldReleaseApplication = shouldReleaseApplication;
            RetainedInstancePresent = retainedInstancePresent;
            Reason = reason ?? string.Empty;
        }

        internal bool DisposalRequired { get; }

        internal bool ShouldQuitApplication { get; }

        internal bool ShouldReleaseApplication { get; }

        internal bool RetainedInstancePresent { get; }

        internal string Reason { get; }

        internal static CaseWorkbookHiddenAppCacheDisposalDecision NotRequired()
        {
            return new CaseWorkbookHiddenAppCacheDisposalDecision(
                disposalRequired: false,
                shouldQuitApplication: false,
                shouldReleaseApplication: false,
                retainedInstancePresent: false,
                reason: "slot-missing");
        }

        internal static CaseWorkbookHiddenAppCacheDisposalDecision SkipNotOwned(bool applicationPresent)
        {
            return new CaseWorkbookHiddenAppCacheDisposalDecision(
                disposalRequired: true,
                shouldQuitApplication: false,
                shouldReleaseApplication: false,
                retainedInstancePresent: applicationPresent,
                reason: "not-cache-owned");
        }

        internal static CaseWorkbookHiddenAppCacheDisposalDecision DisposeOwned(bool applicationPresent)
        {
            return new CaseWorkbookHiddenAppCacheDisposalDecision(
                disposalRequired: true,
                shouldQuitApplication: applicationPresent,
                shouldReleaseApplication: applicationPresent,
                retainedInstancePresent: applicationPresent,
                reason: "cache-owned");
        }
    }
}
