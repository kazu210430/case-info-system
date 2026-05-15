using System;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class UserFacingException : InvalidOperationException
	{
		internal UserFacingException (string userMessage, string diagnosticMessage)
			: base (NormalizeDiagnosticMessage (userMessage, diagnosticMessage))
		{
			UserMessage = NormalizeUserMessage (userMessage);
		}

		internal UserFacingException (string userMessage, string diagnosticMessage, Exception innerException)
			: base (NormalizeDiagnosticMessage (userMessage, diagnosticMessage), innerException)
		{
			UserMessage = NormalizeUserMessage (userMessage);
		}

		internal string UserMessage { get; private set; }

		private static string NormalizeUserMessage (string userMessage)
		{
			return string.IsNullOrWhiteSpace (userMessage) ? "処理結果を確認できませんでした。入力内容をご確認ください。" : userMessage.Trim ();
		}

		private static string NormalizeDiagnosticMessage (string userMessage, string diagnosticMessage)
		{
			return string.IsNullOrWhiteSpace (diagnosticMessage) ? NormalizeUserMessage (userMessage) : diagnosticMessage.Trim ();
		}
	}
}
