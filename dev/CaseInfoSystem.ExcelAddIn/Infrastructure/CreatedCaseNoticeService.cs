using System;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    /// <summary>
    internal sealed class CreatedCaseNoticeService
    {
        private const string NoticeCaption = "案件情報System";
        private const string CreatedMessage = "案件情報Systemを作成しました。";

        private readonly Logger _logger;

        /// <summary>
        internal CreatedCaseNoticeService(Logger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        internal void ShowCreatedCaseCompleted(KernelCaseCreationMode mode)
        {
            MessageBox.Show(
                CreatedMessage,
                NoticeCaption,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            _logger.Info("Created CASE completion notice shown. mode=" + mode.ToString());
        }
    }
}
