using System;
using CaseInfoSystem.WordAddIn.Services;
using Microsoft.Office.Interop.Word;
using Xunit;

namespace CaseInfoSystem.Tests
{
    public class ContentControlBatchReplaceServiceTests
    {
        [Fact]
        public void Execute_WhenExactMatch_ReplacesSupportedControlsOnly()
        {
            var service = new ContentControlBatchReplaceService();
            Document document = new Document();
            ContentControl supported = CreateControl("OldTag", "OldTitle", WdContentControlType.wdContentControlText, start: 5, end: 10);
            ContentControl unsupported = CreateControl("OldTag", "OldTitle", WdContentControlType.wdContentControlCheckBox, start: 20, end: 25);
            document.ContentControls.Add(supported);
            document.ContentControls.Add(unsupported);

            ContentControlBatchReplaceService.ReplaceResult result = service.Execute(document, new ContentControlBatchReplaceService.ReplaceRequest
            {
                OldTag = "OldTag",
                NewTag = "NewTag",
                OldTitle = "OldTitle",
                NewTitle = "NewTitle",
                UsePartialMatch = false
            });

            Assert.Equal(1, result.ScannedCount);
            Assert.Equal(1, result.TagChangedCount);
            Assert.Equal(1, result.TitleChangedCount);
            Assert.Equal("NewTag", supported.Tag);
            Assert.Equal("NewTitle", supported.Title);
            Assert.Equal("OldTag", unsupported.Tag);
            Assert.Equal("OldTitle", unsupported.Title);
        }

        [Fact]
        public void Execute_WhenPartialMatch_ReplacesAllMatchingControlsAndLeavesOthersUntouched()
        {
            var service = new ContentControlBatchReplaceService();
            Document document = new Document();
            ContentControl first = CreateControl("Case-001", "Title-A", WdContentControlType.wdContentControlRichText, start: 1, end: 2);
            ContentControl second = CreateControl("Case-002", "Title-B", WdContentControlType.wdContentControlText, start: 3, end: 4);
            ContentControl third = CreateControl("Other", "NoMatch", WdContentControlType.wdContentControlText, start: 5, end: 6);
            document.ContentControls.Add(first);
            document.ContentControls.Add(second);
            document.ContentControls.Add(third);

            ContentControlBatchReplaceService.ReplaceResult result = service.Execute(document, new ContentControlBatchReplaceService.ReplaceRequest
            {
                OldTag = "Case-",
                NewTag = "Client-",
                UsePartialMatch = true
            });

            Assert.Equal(3, result.ScannedCount);
            Assert.Equal(2, result.TagChangedCount);
            Assert.Equal(0, result.TitleChangedCount);
            Assert.Equal("Client-001", first.Tag);
            Assert.Equal("Client-002", second.Tag);
            Assert.Equal("Other", third.Tag);
        }

        [Fact]
        public void ExecuteNextFromSelection_WhenMatchingControlExistsAfterSelection_ReplacesOneAndMovesSelection()
        {
            var service = new ContentControlBatchReplaceService();
            Document document = new Document();
            ContentControl first = CreateControl("Alpha", "TitleA", WdContentControlType.wdContentControlText, start: 5, end: 8);
            ContentControl second = CreateControl("Beta", "TitleB", WdContentControlType.wdContentControlText, start: 10, end: 20);
            document.ContentControls.Add(first);
            document.ContentControls.Add(second);
            Selection selection = new Selection();
            selection.SetRange(9, 9);

            ContentControlBatchReplaceService.NextReplaceResult result = service.ExecuteNextFromSelection(document, selection, new ContentControlBatchReplaceService.ReplaceRequest
            {
                OldTag = "Beta",
                NewTag = "Gamma"
            });

            Assert.True(result.FoundMatch);
            Assert.Equal(1, result.TagChangedCount);
            Assert.Equal(0, result.TitleChangedCount);
            Assert.Equal(10, result.ControlStart);
            Assert.Equal("Gamma", result.ControlTag);
            Assert.Equal(20, selection.Range.Start);
            Assert.Equal(20, selection.Range.End);
            Assert.Equal("Gamma", second.Tag);
            Assert.Equal("Alpha", first.Tag);
        }

        [Fact]
        public void ExecuteNextFromSelection_WhenNoMatchingControlExists_ReturnsFoundFalse()
        {
            var service = new ContentControlBatchReplaceService();
            Document document = new Document();
            document.ContentControls.Add(CreateControl("Alpha", "TitleA", WdContentControlType.wdContentControlText, start: 5, end: 8));
            Selection selection = new Selection();
            selection.SetRange(9, 9);

            ContentControlBatchReplaceService.NextReplaceResult result = service.ExecuteNextFromSelection(document, selection, new ContentControlBatchReplaceService.ReplaceRequest
            {
                OldTag = "Beta",
                NewTag = "Gamma"
            });

            Assert.False(result.FoundMatch);
            Assert.Equal(9, selection.Range.Start);
            Assert.Equal(9, selection.Range.End);
        }

        [Fact]
        public void Execute_WhenTagWriteFails_PropagatesCurrentExceptionBehavior()
        {
            var service = new ContentControlBatchReplaceService();
            Document document = new Document();
            ContentControl control = CreateControl("Alpha", "TitleA", WdContentControlType.wdContentControlText, start: 1, end: 2);
            control.ThrowOnTagSet = true;
            document.ContentControls.Add(control);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => service.Execute(document, new ContentControlBatchReplaceService.ReplaceRequest
            {
                OldTag = "Alpha",
                NewTag = "Beta"
            }));

            Assert.Contains("tag write failed", exception.Message);
        }

        private static ContentControl CreateControl(string tag, string title, WdContentControlType type, int start, int end)
        {
            ContentControl control = new ContentControl
            {
                Type = type
            };
            control.Tag = tag;
            control.Title = title;
            control.Range.Start = start;
            control.Range.End = end;
            return control;
        }
    }
}
