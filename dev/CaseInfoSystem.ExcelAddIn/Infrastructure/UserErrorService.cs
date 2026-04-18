using System;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class UserErrorService
	{
		private readonly Logger _logger;

		internal UserErrorService (Logger logger)
		{
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void ShowUserError (string procedureName, Exception exception)
		{
			string text = (string.IsNullOrWhiteSpace (procedureName) ? "(unknown)" : procedureName.Trim ());
			Exception ex = exception ?? new InvalidOperationException ("不明なエラーが発生しました。");
			_logger.Error (text, ex);
			string text2 = "エラーが発生しました。" + Environment.NewLine + Environment.NewLine + "処理: " + text + Environment.NewLine + "種類: " + ex.GetType ().Name + Environment.NewLine + "内容: " + ex.Message;
			MessageBox.Show (text2, "案件情報System", MessageBoxButtons.OK, MessageBoxIcon.Hand);
		}
	}
}
