using System;
using System.Windows.Forms;
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
            _contentControlBatchReplaceService = new ContentControlBatchReplaceService();
            ((Word.ApplicationEvents4_Event)Application).DocumentOpen += Application_DocumentOpen;
            ((Word.ApplicationEvents4_Event)Application).NewDocument += Application_NewDocument;
            ((Word.ApplicationEvents4_Event)Application).WindowActivate += Application_WindowActivate;
            ApplyStylesPaneVisibility(Application == null ? null : Application.ActiveDocument);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ((Word.ApplicationEvents4_Event)Application).DocumentOpen -= Application_DocumentOpen;
            ((Word.ApplicationEvents4_Event)Application).NewDocument -= Application_NewDocument;
            ((Word.ApplicationEvents4_Event)Application).WindowActivate -= Application_WindowActivate;
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
            if (Application == null)
            {
                return;
            }

            try
            {
                Application.TaskPanes[Word.WdTaskPanes.wdTaskPaneFormatting].Visible = IsStylePaneEnabled(document);
            }
            catch
            {
            }

            UpdateRibbonToggleState(document);
        }

        internal void RegisterRibbon(WordApplicationRibbon ribbon)
        {
            _wordApplicationRibbon = ribbon;
            UpdateRibbonToggleState(Application == null ? null : Application.ActiveDocument);
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
            _wordApplicationRibbon = new WordApplicationRibbon();

            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new Microsoft.Office.Tools.Ribbon.IRibbonExtension[]
            {
                _wordApplicationRibbon
            });
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

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}
