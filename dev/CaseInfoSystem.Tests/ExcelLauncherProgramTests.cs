using System;
using System.IO;
using System.Reflection;
using CaseInfoSystem.ExcelLauncher;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class ExcelLauncherProgramTests
    {
        [Fact]
        public void KernelWorkbookFileName_IsExpectedWorkbookName()
        {
            FieldInfo field = typeof(Program).GetField("KernelWorkbookFileName", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.NotNull(field);
            Assert.Equal("案件情報System_Kernel.xlsx", field.GetRawConstantValue());
        }

        [Fact]
        public void ResolveKernelWorkbookPath_UsesKernelWorkbookFileNameConstant()
        {
            string baseDirectory = Path.Combine("C:\\launcher", "bin", "Debug");
            string workbookFileName = (string)typeof(Program)
                .GetField("KernelWorkbookFileName", BindingFlags.NonPublic | BindingFlags.Static)
                .GetRawConstantValue();

            string workbookPath = Program.ResolveKernelWorkbookPath(baseDirectory);

            Assert.Equal(
                Path.Combine(baseDirectory, workbookFileName),
                workbookPath);
            Assert.Equal(workbookFileName, Path.GetFileName(workbookPath));
        }

        [Fact]
        public void ValidateFileExists_WhenFileIsMissing_ThrowsExpectedMessage()
        {
            string missingPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "missing.xlsx");
            MethodInfo method = typeof(Program).GetMethod("ValidateFileExists", BindingFlags.NonPublic | BindingFlags.Static);

            TargetInvocationException exception = Assert.Throws<TargetInvocationException>(
                () => method.Invoke(null, new object[] { missingPath }));

            FileNotFoundException innerException = Assert.IsType<FileNotFoundException>(exception.InnerException);
            Assert.StartsWith("起動対象ファイルが存在しません。", innerException.Message);
            Assert.Equal(missingPath, innerException.FileName);
        }
    }
}
