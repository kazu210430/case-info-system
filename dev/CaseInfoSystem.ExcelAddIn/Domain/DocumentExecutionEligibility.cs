namespace CaseInfoSystem.ExcelAddIn.Domain
{
    /// <summary>
    internal sealed class DocumentExecutionEligibility
    {
        /// <summary>
        internal DocumentExecutionEligibility(bool canExecuteInVsto, string reason, DocumentTemplateSpec templateSpec, CaseContext caseContext = null)
        {
            CanExecuteInVsto = canExecuteInVsto;
            Reason = reason ?? string.Empty;
            TemplateSpec = templateSpec;
            CaseContext = caseContext;
        }

        /// <summary>
        internal bool CanExecuteInVsto { get; }

        /// <summary>
        internal string Reason { get; }

        /// <summary>
        internal DocumentTemplateSpec TemplateSpec { get; }

        /// <summary>
        internal CaseContext CaseContext { get; }

        /// <summary>
        internal DocumentTemplateResolutionSource ResolutionSource
        {
            get { return TemplateSpec == null ? DocumentTemplateResolutionSource.Unknown : TemplateSpec.ResolutionSource; }
        }
    }
}
