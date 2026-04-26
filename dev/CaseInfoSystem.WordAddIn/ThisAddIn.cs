using System;
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

        private ContentControlBatchReplaceService _contentControlBatchReplaceService;
        private WordApplicationRibbon _wordApplicationRibbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            WordAddInStartupLogWriter.Write("WordAddin startup begin");

            try
            {
                ExecuteStartupStep("service composition", delegate
                {
                    _contentControlBatchReplaceService = new ContentControlBatchReplaceService();
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
                    // Document-specific pane state is refreshed by the existing document lifecycle events.
                    ApplyStylesPaneVisibility(null, true, "taskpane or pane initialization");
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
                Word.ApplicationEvents4_Event applicationEvents = (Word.ApplicationEvents4_Event)Application;
                applicationEvents.DocumentOpen -= Application_DocumentOpen;
                applicationEvents.NewDocument -= Application_NewDocument;
                applicationEvents.WindowActivate -= Application_WindowActivate;
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
            ApplyStylesPaneVisibility(document);
        }

        private void Application_NewDocument(Word.Document document)
        {
            ApplyStylesPaneVisibility(document);
        }

        private void Application_WindowActivate(Word.Document document, Word.Window window)
        {
            ApplyStylesPaneVisibility(document);
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
            ApplyStylesPaneVisibility(activeDocument);
        }

        private void ApplyStylesPaneVisibility(Word.Document document)
        {
            ApplyStylesPaneVisibility(document, false, null);
        }

        private void ApplyStylesPaneVisibility(Word.Document document, bool throwOnFailure, string diagnosticStepName)
        {
            if (Application == null)
            {
                if (!string.IsNullOrWhiteSpace(diagnosticStepName))
                {
                    WordAddInStartupLogWriter.Write(diagnosticStepName + " skipped because Application is null");
                }

                return;
            }

            try
            {
                Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneFormatting].Visible = IsStylePaneEnabled(document);
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
            }

            UpdateRibbonToggleState(document);
        }

        internal void RegisterRibbon(WordApplicationRibbon ribbon)
        {
            _wordApplicationRibbon = ribbon;
            UpdateRibbonToggleState(TryGetActiveDocument("RegisterRibbon initial state"));
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
            catch
            {
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
            catch
            {
            }

            return IntPtr.Zero;
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
