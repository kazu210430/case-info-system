using System;
using System.Collections.Generic;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
	internal sealed class KernelTemplateSyncPreflightRequest
	{
		internal KernelTemplateSyncPreflightRequest (string systemRoot, IReadOnlyCollection<string> definedTemplateTags)
		{
			SystemRoot = systemRoot ?? string.Empty;
			DefinedTemplateTags = definedTemplateTags ?? Array.Empty<string> ();
		}

		internal string SystemRoot { get; }

		internal IReadOnlyCollection<string> DefinedTemplateTags { get; }
	}
}
