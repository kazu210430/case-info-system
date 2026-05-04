using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
    public class ExcelApplicationStateScopeTests
    {
        [Fact]
        public void Dispose_RestoresOnlyTouchedStates()
        {
            var application = new Excel.Application
            {
                ScreenUpdating = true,
                EnableEvents = true,
                DisplayAlerts = true,
                Calculation = Excel.XlCalculation.xlCalculationAutomatic
            };

            using (var scope = new ExcelApplicationStateScope(application))
            {
                scope.SetScreenUpdating(false);
                scope.SetEnableEvents(false);
            }

            Assert.True(application.ScreenUpdating);
            Assert.True(application.EnableEvents);
            Assert.True(application.DisplayAlerts);
            Assert.Equal(Excel.XlCalculation.xlCalculationAutomatic, application.Calculation);
        }

        [Fact]
        public void Dispose_RestoresCalculationWhenExplicitlyChanged()
        {
            var application = new Excel.Application
            {
                ScreenUpdating = true,
                EnableEvents = true,
                DisplayAlerts = true,
                Calculation = Excel.XlCalculation.xlCalculationAutomatic
            };

            using (var scope = new ExcelApplicationStateScope(application))
            {
                scope.SetCalculation(Excel.XlCalculation.xlCalculationManual);
                Assert.Equal(Excel.XlCalculation.xlCalculationManual, application.Calculation);
            }

            Assert.Equal(Excel.XlCalculation.xlCalculationAutomatic, application.Calculation);
        }

        [Fact]
        public void NestedScopes_RestoreToOuterScopeState()
        {
            var application = new Excel.Application
            {
                ScreenUpdating = true,
                EnableEvents = true,
                DisplayAlerts = true,
                Calculation = Excel.XlCalculation.xlCalculationAutomatic
            };

            using (var outerScope = new ExcelApplicationStateScope(application))
            {
                outerScope.SetScreenUpdating(false);
                outerScope.SetDisplayAlerts(false);

                using (var innerScope = new ExcelApplicationStateScope(application))
                {
                    innerScope.SetScreenUpdating(true);
                    innerScope.SetDisplayAlerts(true);
                }

                Assert.False(application.ScreenUpdating);
                Assert.False(application.DisplayAlerts);
            }

            Assert.True(application.ScreenUpdating);
            Assert.True(application.DisplayAlerts);
        }

        [Fact]
        public void NullApplication_IsSafe()
        {
            using (var scope = new ExcelApplicationStateScope(null))
            {
                scope.SetScreenUpdating(false);
                scope.SetEnableEvents(false);
                scope.SetDisplayAlerts(false);
                scope.SetCalculation(Excel.XlCalculation.xlCalculationManual);
            }
        }
    }
}
