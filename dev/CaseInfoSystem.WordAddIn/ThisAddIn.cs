using System;
using System.IO;
using System.Windows.Forms;
using CaseInfoSystem.WordAddIn.Infrastructure;
using CaseInfoSystem.WordAddIn.Services;
using CaseInfoSystem.WordAddIn.UI;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using RibbonExtensibility = Microsoft.Office.Core.IRibbonExtensibility;

namespace CaseInfoSystem.WordAddIn
{
    public partial class ThisAddIn
    {
        private const string StylePaneEnabledPropertyName = "StylePaneEnabled";
        private const int StylePaneActivationRetryIntervalMilliseconds = 250;
        private const int StylePaneActivationRetryMaxAttempts = 20;

        private ContentControlBatchReplaceService _contentControlBatchReplaceService;
        private ContentControlFolderBatchReplaceService _contentControlFolderBatchReplaceService;
        private WordApplicationRibbon _wordApplicationRibbon;
        private Timer _stylePaneActivationRetryTimer;
        private int _stylePaneActivationRetryAttemptsRemaining;
        private string _stylePaneActivationRetryReason;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            WordAddInStartupLogWriter.Write("WordAddin startup begin");

            try
            {
                ExecuteStartupStep("service composition", delegate
                {
                    _contentControlBatchReplaceService = new ContentControlBatchReplaceService();
                    _contentControlFolderBatchReplaceService = new ContentControlFolderBatchReplaceService();
                    _stylePaneActivationRetryTimer = new Timer
                    {
                        Interval = StylePaneActivationRetryIntervalMilliseconds
                    };
                    _stylePaneActivationRetryTimer.Tick += StylePaneActivationRetryTimer_Tick;
                });

                ExecuteStartupStep("event subscribe", delegate
                {
                    Word.ApplicationEvents4_Event applicationEvents = (Word.ApplicationEvents4_Event)Application;
                    applicationEvents.DocumentOpen += Application_DocumentOpen;
                    applicationEvents.NewDocument += Application_NewDocument;
                    applicationEvents.WindowActivate += Application_WindowActivate;
                });

                ExecuteStartupStep("taskpane or pane initialization", delegate
                {
                    // ActiveDocument is not guaranteed to exist during startup.
                    // ApplyStylePaneVisibilityWithRetry schedules a bounded retry so the style pane can be re-applied after Word finishes preparing the window/document.
                    ApplyStylePaneVisibilityWithRetry(TryGetActiveDocument("taskpane or pane initialization"), "taskpane or pane initialization");
                });

                WordAddInStartupLogWriter.Write("WordAddin startup end");
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("WordAddin startup failure", ex);
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            WordAddInStartupLogWriter.Write("WordAddin shutdown begin");

            try
            {
                StopStylePaneActivationRetry("shutdown");

                Word.ApplicationEvents4_Event applicationEvents = (Word.ApplicationEvents4_Event)Application;
                applicationEvents.DocumentOpen -= Application_DocumentOpen;
                applicationEvents.NewDocument -= Application_NewDocument;
                applicationEvents.WindowActivate -= Application_WindowActivate;
                if (_stylePaneActivationRetryTimer != null)
                {
                    _stylePaneActivationRetryTimer.Tick -= StylePaneActivationRetryTimer_Tick;
                    _stylePaneActivationRetryTimer.Dispose();
                    _stylePaneActivationRetryTimer = null;
                }

                WordAddInStartupLogWriter.Write("WordAddin shutdown end");
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("WordAddin shutdown failure", ex);
                throw;
            }
        }

        private void Application_DocumentOpen(Word.Document document)
        {
            ApplyStylePaneVisibilityWithRetry(document, "DocumentOpen");
        }

        private void Application_NewDocument(Word.Document document)
        {
            ApplyStylePaneVisibilityWithRetry(document, "NewDocument");
        }

        private void Application_WindowActivate(Word.Document document, Word.Window window)
        {
            ApplyStylePaneVisibilityWithRetry(document, "WindowActivate");
        }

        internal void ToggleStylePaneForActiveDocument()
        {
            Word.Document activeDocument = Application == null ? null : Application.ActiveDocument;
            if (activeDocument == null)
            {
                UpdateRibbonToggleState(null);
                MessageBox.Show("文書を開いてから実行してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool nextState = !IsStylePaneEnabled(activeDocument);
            SetStylePaneEnabled(activeDocument, nextState);
            ApplyStylePaneVisibilityWithRetry(activeDocument, "ToggleStylePaneForActiveDocument");
        }

        private void ApplyStylePaneVisibilityWithRetry(Word.Document document, string reason)
        {
            if (ApplyStylesPaneVisibility(document, false, reason))
            {
                StopStylePaneActivationRetry(reason + " immediate success");
                return;
            }

            ScheduleStylePaneActivationRetry(reason);
        }

        private bool ApplyStylesPaneVisibility(Word.Document document, bool throwOnFailure, string diagnosticStepName)
        {
            if (Application == null)
            {
                if (!string.IsNullOrWhiteSpace(diagnosticStepName))
                {
                    WordAddInStartupLogWriter.Write(diagnosticStepName + " skipped because Application is null");
                }

                return false;
            }

            if (document == null)
            {
                if (!string.IsNullOrWhiteSpace(diagnosticStepName))
                {
                    WordAddInStartupLogWriter.Write(diagnosticStepName + " skipped because document is unavailable");
                }

                UpdateRibbonToggleState(null);
                return false;
            }

            bool stylePaneEnabled = IsStylePaneEnabled(document);

            try
            {
                Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneFormatting].Visible = stylePaneEnabled;

                if (!string.IsNullOrWhiteSpace(diagnosticStepName))
                {
                    WordAddInStartupLogWriter.Write(
                        diagnosticStepName
                        + " applied style pane visibility="
                        + stylePaneEnabled
                        + " document="
                        + DescribeDocument(document));
                }
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrWhiteSpace(diagnosticStepName))
                {
                    WordAddInStartupLogWriter.WriteException(diagnosticStepName + " failure", ex);
                }

                if (throwOnFailure)
                {
                    throw;
                }

                UpdateRibbonToggleState(document);
                return false;
            }

            UpdateRibbonToggleState(document);
            return true;
        }

        internal void RegisterRibbon(WordApplicationRibbon ribbon)
        {
            _wordApplicationRibbon = ribbon;
            UpdateRibbonToggleState(TryGetActiveDocument("RegisterRibbon initial state"));
            ScheduleStylePaneActivationRetry("RegisterRibbon");
        }

        private void UpdateRibbonToggleState(Word.Document document)
        {
            if (_wordApplicationRibbon == null)
            {
                return;
            }

            _wordApplicationRibbon.SyncStylePaneToggleState(document != null && IsStylePaneEnabled(document), document != null);
        }

        private static bool IsStylePaneEnabled(Word.Document document)
        {
            if (document == null)
            {
                return false;
            }

            Office.DocumentProperties properties = null;

            try
            {
                properties = (Office.DocumentProperties)document.CustomDocumentProperties;
                foreach (Office.DocumentProperty property in properties)
                {
                    if (!string.Equals(property.Name, StylePaneEnabledPropertyName, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    object value = property.Value;
                    if (value is bool boolValue)
                    {
                        return boolValue;
                    }

                    bool parsedValue;
                    return bool.TryParse(Convert.ToString(value), out parsedValue) && parsedValue;
                }
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("IsStylePaneEnabled failure for " + DescribeDocument(document), ex);
            }

            return false;
        }

        private static void SetStylePaneEnabled(Word.Document document, bool enabled)
        {
            if (document == null)
            {
                return;
            }

            Office.DocumentProperties properties = (Office.DocumentProperties)document.CustomDocumentProperties;
            foreach (Office.DocumentProperty property in properties)
            {
                if (!string.Equals(property.Name, StylePaneEnabledPropertyName, StringComparison.Ordinal))
                {
                    continue;
                }

                property.Value = enabled;
                return;
            }

            properties.Add(
                StylePaneEnabledPropertyName,
                false,
                Office.MsoDocProperties.msoPropertyTypeBoolean,
                enabled);
        }

        public void ShowContentControlBatchReplaceForm()
        {
            Word.Document activeDocument = Application == null ? null : Application.ActiveDocument;
            if (activeDocument == null)
            {
                MessageBox.Show("アクティブな文書がありません。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var form = new ContentControlBatchReplaceForm())
            {
                var owner = new WordWindowOwner(ResolveWordWindowHandle());
                if (form.ShowDialog(owner) != DialogResult.OK)
                {
                    return;
                }

                var result = _contentControlBatchReplaceService.Execute(activeDocument, form.ReplaceRequest);
                MessageBox.Show(
                    ContentControlBatchReplaceService.BuildCompletionMessage(result),
                    "案件情報System",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        public void ShowContentControlFolderBatchReplaceForm()
        {
            if (_contentControlFolderBatchReplaceService == null)
            {
                MessageBox.Show("雛形フォルダ一括置換サービスを利用できません。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var form = new ContentControlFolderBatchReplaceForm(ResolveInitialTemplateDirectory()))
            {
                var owner = new WordWindowOwner(ResolveWordWindowHandle());
                if (form.ShowDialog(owner) != DialogResult.OK)
                {
                    return;
                }

                DialogResult confirmation = MessageBox.Show(
                    "雛形フォルダ直下の対象ファイルを直接更新します。実行しますか？",
                    "案件情報System",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2);
                if (confirmation != DialogResult.Yes)
                {
                    return;
                }

                try
                {
                    ContentControlFolderBatchReplaceService.FolderReplaceResult result = _contentControlFolderBatchReplaceService.Execute(form.ReplaceRequest);
                    MessageBox.Show(
                        ContentControlFolderBatchReplaceService.BuildCompletionMessage(result),
                        "案件情報System",
                        MessageBoxButtons.OK,
                        result.FailedFileCount == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    WordAddInStartupLogWriter.WriteException("ShowContentControlFolderBatchReplaceForm failure", ex);
                    MessageBox.Show("雛形フォルダ一括置換を実行できませんでした。ログを確認してください。", "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        internal ContentControlBatchReplaceService.NextReplaceResult ReplaceNextContentControlFromSelection(ContentControlBatchReplaceService.ReplaceRequest request)
        {
            Word.Document activeDocument = Application == null ? null : Application.ActiveDocument;
            if (activeDocument == null)
            {
                throw new InvalidOperationException("アクティブな文書がありません。");
            }

            Word.Selection activeSelection = Application.Selection;
            if (activeSelection == null)
            {
                throw new InvalidOperationException("選択位置を取得できませんでした。");
            }

            return _contentControlBatchReplaceService.ExecuteNextFromSelection(activeDocument, activeSelection, request);
        }

        protected override RibbonExtensibility CreateRibbonExtensibilityObject()
        {
            WordAddInStartupLogWriter.Write("CreateRibbonExtensibilityObject begin");

            try
            {
                _wordApplicationRibbon = new WordApplicationRibbon();
                RibbonExtensibility ribbonManager = Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[]
                {
                    _wordApplicationRibbon
                });

                WordAddInStartupLogWriter.Write("CreateRibbonExtensibilityObject success");
                return ribbonManager;
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("CreateRibbonExtensibilityObject failure", ex);
                throw;
            }
        }

        private IntPtr ResolveWordWindowHandle()
        {
            try
            {
                Word.Window activeWindow = Application == null ? null : Application.ActiveWindow;
                if (activeWindow != null)
                {
                    return new IntPtr(activeWindow.Hwnd);
                }
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("ResolveWordWindowHandle failure", ex);
            }

            return IntPtr.Zero;
        }

        private string ResolveInitialTemplateDirectory()
        {
            try
            {
                Word.Document activeDocument = Application == null ? null : Application.ActiveDocument;
                string fullName = activeDocument == null ? string.Empty : activeDocument.FullName;
                if (!string.IsNullOrWhiteSpace(fullName) && File.Exists(fullName))
                {
                    return Path.GetDirectoryName(fullName) ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("ResolveInitialTemplateDirectory failure", ex);
            }

            return string.Empty;
        }

        private Word.Document TryGetActiveDocument(string diagnosticContext)
        {
            if (Application == null)
            {
                return null;
            }

            try
            {
                return Application.ActiveDocument;
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException(diagnosticContext + " active document unavailable", ex);
                return null;
            }
        }

        private void ScheduleStylePaneActivationRetry(string reason)
        {
            if (_stylePaneActivationRetryTimer == null)
            {
                WordAddInStartupLogWriter.Write("style pane activation retry skipped because timer is unavailable");
                return;
            }

            _stylePaneActivationRetryAttemptsRemaining = Math.Max(_stylePaneActivationRetryAttemptsRemaining, StylePaneActivationRetryMaxAttempts);
            _stylePaneActivationRetryReason = reason ?? "unknown";

            if (_stylePaneActivationRetryTimer.Enabled)
            {
                WordAddInStartupLogWriter.Write(
                    "style pane activation retry refreshed"
                    + " reason=" + _stylePaneActivationRetryReason
                    + " attemptsRemaining=" + _stylePaneActivationRetryAttemptsRemaining);
                return;
            }

            _stylePaneActivationRetryTimer.Start();
            WordAddInStartupLogWriter.Write(
                "style pane activation retry scheduled"
                + " reason=" + _stylePaneActivationRetryReason
                + " intervalMs=" + StylePaneActivationRetryIntervalMilliseconds
                + " attemptsRemaining=" + _stylePaneActivationRetryAttemptsRemaining);
        }

        private void StopStylePaneActivationRetry(string reason)
        {
            if (_stylePaneActivationRetryTimer == null || !_stylePaneActivationRetryTimer.Enabled)
            {
                return;
            }

            _stylePaneActivationRetryTimer.Stop();
            _stylePaneActivationRetryAttemptsRemaining = 0;
            _stylePaneActivationRetryReason = null;

            if (!string.IsNullOrWhiteSpace(reason))
            {
                WordAddInStartupLogWriter.Write("style pane activation retry stopped reason=" + reason);
            }
        }

        private void StylePaneActivationRetryTimer_Tick(object sender, EventArgs e)
        {
            string reason = _stylePaneActivationRetryReason ?? "timer";
            string diagnosticContext = reason + " retry attempt " + (StylePaneActivationRetryMaxAttempts - _stylePaneActivationRetryAttemptsRemaining + 1);
            Word.Document activeDocument = TryGetActiveDocument(diagnosticContext);
            if (ApplyStylesPaneVisibility(activeDocument, false, diagnosticContext))
            {
                StopStylePaneActivationRetry(diagnosticContext + " success");
                return;
            }

            _stylePaneActivationRetryAttemptsRemaining--;
            if (_stylePaneActivationRetryAttemptsRemaining > 0)
            {
                WordAddInStartupLogWriter.Write(
                    diagnosticContext
                    + " pending"
                    + " attemptsRemaining=" + _stylePaneActivationRetryAttemptsRemaining);
                return;
            }

            StopStylePaneActivationRetry(diagnosticContext + " exhausted");
        }

        private static string DescribeDocument(Word.Document document)
        {
            if (document == null)
            {
                return "<null>";
            }

            try
            {
                return string.IsNullOrWhiteSpace(document.FullName) ? document.Name : document.FullName;
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("DescribeDocument failure", ex);
                return "<unavailable>";
            }
        }

        private void InternalStartup()
        {
            WordAddInStartupLogWriter.Write("InternalStartup begin");

            try
            {
                Startup += ThisAddIn_Startup;
                Shutdown += ThisAddIn_Shutdown;
                WordAddInStartupLogWriter.Write("InternalStartup success");
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException("InternalStartup failure", ex);
                throw;
            }
        }

        private static void ExecuteStartupStep(string stepName, Action action)
        {
            WordAddInStartupLogWriter.Write(stepName + " begin");

            try
            {
                action();
                WordAddInStartupLogWriter.Write(stepName + " success");
            }
            catch (Exception ex)
            {
                WordAddInStartupLogWriter.WriteException(stepName + " failure", ex);
                throw;
            }
        }
    }
}
