using System;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class KernelTemplateSyncPreflightService
	{
		private const string TemplateFolderName = "雛形";

		private const string MissingDefinedTemplateTagsMessage = "Kernelブックの管理シート CaseList_FieldInventory を読み取れません。";

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly WordTemplateRegistrationValidationService _wordTemplateRegistrationValidationService;

		internal KernelTemplateSyncPreflightService (PathCompatibilityService pathCompatibilityService, WordTemplateRegistrationValidationService wordTemplateRegistrationValidationService)
		{
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_wordTemplateRegistrationValidationService = wordTemplateRegistrationValidationService ?? throw new ArgumentNullException ("wordTemplateRegistrationValidationService");
		}

		internal KernelTemplateSyncPreflightResult Run (KernelTemplateSyncPreflightRequest request)
		{
			if (request == null) {
				throw new ArgumentNullException ("request");
			}
			string text = ResolveTemplateDirectory (request.SystemRoot);
			if (request.DefinedTemplateTags == null || request.DefinedTemplateTags.Count == 0) {
				return KernelTemplateSyncPreflightResult.Failed (text, new ValidationFailureSummary (ValidationFailureKind.MissingDefinedTemplateTags, MissingDefinedTemplateTagsMessage, 0, Array.Empty<TemplateRegistrationValidationEntry> ()));
			}
			TemplateRegistrationValidationSummary templateRegistrationValidationSummary = _wordTemplateRegistrationValidationService.Validate (text, request.DefinedTemplateTags);
			if (templateRegistrationValidationSummary.DetectedFileCount == 0) {
				return KernelTemplateSyncPreflightResult.Failed (text, new ValidationFailureSummary (ValidationFailureKind.NoTemplateFiles, "雛形フォルダに Word 雛形 (.docx / .dotx / .docm / .dotm) が見つかりませんでした。" + Environment.NewLine + "フォルダ: " + text, 0, templateRegistrationValidationSummary.TemplateResults), templateRegistrationValidationSummary);
			}
			return KernelTemplateSyncPreflightResult.Succeeded (text, templateRegistrationValidationSummary);
		}

		private string ResolveTemplateDirectory (string systemRoot)
		{
			string left = _pathCompatibilityService.NormalizePath (systemRoot);
			string text = _pathCompatibilityService.CombinePath (left, TemplateFolderName);
			if (!_pathCompatibilityService.DirectoryExistsSafe (text)) {
				throw new InvalidOperationException ("雛形フォルダが見つかりません: " + text);
			}
			return text;
		}
	}
}
