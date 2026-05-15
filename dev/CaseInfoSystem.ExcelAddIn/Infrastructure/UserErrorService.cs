using System;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class UserErrorService
	{
		private const string DefaultTitle = "案件情報System";

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
			UserFacingException userFacingException = ex as UserFacingException;
			if (userFacingException != null) {
				ShowOkNotification (userFacingException.UserMessage, DefaultTitle, MessageBoxIcon.Hand);
				return;
			}
			string text2 = "エラーが発生しました。" + Environment.NewLine + Environment.NewLine + "処理: " + text + Environment.NewLine + "種類: " + ex.GetType ().Name + Environment.NewLine + "内容: " + ex.Message;
			ShowOkNotification (text2, DefaultTitle, MessageBoxIcon.Hand);
		}

		internal static DialogResult ShowOkNotification (string message, string title, MessageBoxIcon icon)
		{
			return MessageBox.Show (
				message ?? string.Empty,
				string.IsNullOrWhiteSpace (title) ? DefaultTitle : title,
				MessageBoxButtons.OK,
				icon);
		}
	}
}
