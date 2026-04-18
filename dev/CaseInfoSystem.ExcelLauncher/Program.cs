using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelLauncher
{
    /// <summary>
    /// クラス: Kernel workbook を既定の Excel で起動するランチャー。
    /// 責務: 配布先フォルダにある Kernel workbook の存在確認と起動だけを行う。
    /// </summary>
    internal static class Program
    {
        private const string KernelWorkbookFileName = "案件情報System_Kernel.xlsx";

        /// <summary>
        /// メソッド: ランチャーのエントリポイント。
        /// 引数: args - 起動引数。
        /// 戻り値: 正常終了時は 0、エラー時は非 0。
        /// 副作用: Kernel workbook を既定の Excel で起動する。
        /// </summary>
        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                string workbookPath = ResolveKernelWorkbookPath();
                ValidateFileExists(workbookPath);
                LaunchWorkbookWithShell(workbookPath);
                return 0;
            }
            catch (FileNotFoundException fileNotFoundException)
            {
                ShowErrorMessage(
                    "起動対象の Excel ファイルが見つかりませんでした。EXE と同じフォルダに対象ファイルがあることを確認してください。"
                    + Environment.NewLine
                    + Environment.NewLine
                    + fileNotFoundException.FileName,
                    "ファイル未検出");
                return 1;
            }
            catch (Exception exception)
            {
                ShowErrorMessage(
                    "ランチャー起動中に予期しないエラーが発生しました。"
                    + Environment.NewLine
                    + Environment.NewLine
                    + exception.Message,
                    "起動エラー");
                return 2;
            }
        }

        /// <summary>
        /// メソッド: 実行フォルダ基準で Kernel workbook パスを解決する。
        /// 引数: なし。
        /// 戻り値: Kernel workbook のフルパス。
        /// 副作用: なし。
        /// </summary>
        private static string ResolveKernelWorkbookPath()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(baseDirectory, KernelWorkbookFileName);
        }

        /// <summary>
        /// メソッド: 対象 workbook の存在を確認する。
        /// 引数: workbookPath - 確認対象ファイルパス。
        /// 戻り値: なし。
        /// 副作用: ファイルがない場合は FileNotFoundException を送出する。
        /// </summary>
        private static void ValidateFileExists(string workbookPath)
        {
            if (!File.Exists(workbookPath))
            {
                throw new FileNotFoundException("起動対象ファイルが存在しません。", workbookPath);
            }
        }

        /// <summary>
        /// メソッド: 対象 workbook を既定アプリで起動する。
        /// 引数: workbookPath - 起動対象ファイルパス。
        /// 戻り値: なし。
        /// 副作用: Excel プロセス起動を OS に委譲する。
        /// </summary>
        private static void LaunchWorkbookWithShell(string workbookPath)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = workbookPath,
                UseShellExecute = true,
                WorkingDirectory = Path.GetDirectoryName(workbookPath) ?? AppDomain.CurrentDomain.BaseDirectory
            };

            Process.Start(startInfo);
        }

        /// <summary>
        /// メソッド: エラーメッセージを表示する。
        /// 引数: message - 表示文言, title - ダイアログタイトル。
        /// 戻り値: なし。
        /// 副作用: モーダルダイアログを表示する。
        /// </summary>
        private static void ShowErrorMessage(string message, string title)
        {
            MessageBox.Show(
                message,
                title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1);
        }
    }
}
