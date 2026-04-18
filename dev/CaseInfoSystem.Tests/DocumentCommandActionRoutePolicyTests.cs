using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class DocumentCommandActionRoutePolicyTests
    {
        [Theory]
        [InlineData("doc")]
        [InlineData("DOC")]
        public void Decide_ReturnsDocument_ForDocumentAction(string actionKind)
        {
            DocumentCommandActionRoute route = DocumentCommandActionRoutePolicy.Decide(actionKind);

            Assert.Equal(DocumentCommandActionRoute.Document, route);
        }

        [Fact]
        public void Decide_ReturnsAccounting_ForAccountingAction()
        {
            DocumentCommandActionRoute route = DocumentCommandActionRoutePolicy.Decide("accounting");

            Assert.Equal(DocumentCommandActionRoute.Accounting, route);
        }

        [Fact]
        public void Decide_ReturnsCaseList_ForCaseListAction()
        {
            DocumentCommandActionRoute route = DocumentCommandActionRoutePolicy.Decide("caselist");

            Assert.Equal(DocumentCommandActionRoute.CaseList, route);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("preview")]
        public void Decide_ReturnsUnsupported_ForUnknownAction(string actionKind)
        {
            DocumentCommandActionRoute route = DocumentCommandActionRoutePolicy.Decide(actionKind);

            Assert.Equal(DocumentCommandActionRoute.Unsupported, route);
        }
    }
}
