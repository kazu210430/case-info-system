using System;
using System.Drawing;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.App;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    /// <summary>
    /// クラス: CASE 用 Task Pane の UI コンテナを表す。
    /// 責務: 外部から受け取った描画状態を内側の描画コントロールへ橋渡しする。
    /// </summary>
    internal sealed class DocumentButtonsControl : UserControl, ITaskPaneView
    {
        private readonly DocTaskPaneControl _innerControl;

        internal DocumentButtonsControl()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.MintCream;
            _innerControl = new DocTaskPaneControl
            {
                Dock = DockStyle.Fill
            };
            Controls.Add(_innerControl);
        }

        internal event EventHandler<TaskPaneActionEventArgs> ActionInvoked
        {
            add { _innerControl.ActionInvoked += value; }
            remove { _innerControl.ActionInvoked -= value; }
        }

        public int PreferredWidth
        {
            get { return _innerControl.PreferredPaneWidthHint > 0 ? _innerControl.PreferredPaneWidthHint : 320; }
        }

        internal string SelectedTabName
        {
            get { return _innerControl.SelectedTabName; }
        }

        /// <summary>
        /// CASE pane では Render(viewState) が正規の表示経路。
        /// この API は旧来の直 UI 表示受け口として残置しており、
        /// 新規の CASE メッセージ表示では使わない想定。
        /// </summary>
        internal void ShowMessage(string message)
        {
            _innerControl.ShowMessage(message);
        }

        internal void Render(CaseTaskPaneViewState viewState)
        {
            _innerControl.Render(viewState);
        }
    }
}
