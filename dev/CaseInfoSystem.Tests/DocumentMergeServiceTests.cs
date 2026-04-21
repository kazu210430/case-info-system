using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class DocumentMergeServiceTests
	{
		[Fact]
		public void ApplyMergeData_WhenContentControlMetadataCanOnlyBeReadOnce_UsesCachedValues ()
		{
			List<string> logs = new List<string> ();
			Logger logger = OrchestrationTestSupport.CreateLogger (logs);
			var service = new DocumentMergeService (logger);
			var control = new ReadOnceContentControl ("CustomerName", "CustomerName", type: 1);
			var document = new FakeWordDocument (control);

			service.ApplyMergeData (document, new Dictionary<string, string> { ["CustomerName"] = "Alpha\r\nBeta" });

			Assert.Equal (1, control.TagReadCount);
			Assert.Equal (1, control.TitleReadCount);
			Assert.Equal (1, control.TypeReadCount);
			Assert.Equal ("Alpha\vBeta", control.Range.Text);
		}

		[Fact]
		public void ApplyMergeData_WhenRangeWriteThrows_LogsWarningAndContinues ()
		{
			List<string> logs = new List<string> ();
			Logger logger = OrchestrationTestSupport.CreateLogger (logs);
			var service = new DocumentMergeService (logger);
			var control = new ReadOnceContentControl ("CustomerName", "CustomerName", type: 1, throwOnWrite: true);
			var document = new FakeWordDocument (control);

			Exception exception = null;
			try {
				service.ApplyMergeData (document, new Dictionary<string, string> { ["CustomerName"] = "Alpha" });
			} catch (Exception ex) {
				exception = ex;
			}

			Assert.Null (exception);
			Assert.Contains (logs, message => message.Contains ("DocumentMergeService failed to write content control."));
			Assert.Contains (logs, message => message.Contains ("error=write failed"));
		}

		public sealed class FakeWordDocument
		{
			internal FakeWordDocument (params ReadOnceContentControl[] controls)
			{
				ContentControls = new FakeContentControls (controls);
			}

			public FakeContentControls ContentControls { get; }
		}

		public sealed class FakeContentControls
		{
			private readonly List<ReadOnceContentControl> _controls;

			internal FakeContentControls (IEnumerable<ReadOnceContentControl> controls)
			{
				_controls = new List<ReadOnceContentControl> (controls ?? Array.Empty<ReadOnceContentControl> ());
			}

			public int Count => _controls.Count;

			public ReadOnceContentControl Item (int index)
			{
				return _controls[index - 1];
			}
		}

		public sealed class ReadOnceContentControl
		{
			private readonly string _tag;
			private readonly string _title;
			private readonly int _type;

			internal ReadOnceContentControl (string tag, string title, int type, bool throwOnWrite = false)
			{
				_tag = tag;
				_title = title;
				_type = type;
				Range = new FakeRange (throwOnWrite);
			}

			public int TagReadCount { get; private set; }

			public int TitleReadCount { get; private set; }

			public int TypeReadCount { get; private set; }

			public string Tag
			{
				get
				{
					TagReadCount++;
					if (TagReadCount > 1) {
						throw new InvalidOperationException ("tag reread");
					}
					return _tag;
				}
			}

			public string Title
			{
				get
				{
					TitleReadCount++;
					if (TitleReadCount > 1) {
						throw new InvalidOperationException ("title reread");
					}
					return _title;
				}
			}

			public int Type
			{
				get
				{
					TypeReadCount++;
					if (TypeReadCount > 1) {
						throw new InvalidOperationException ("type reread");
					}
					return _type;
				}
			}

			public FakeRange Range { get; }
		}

		public sealed class FakeRange
		{
			private readonly bool _throwOnWrite;
			private string _text = string.Empty;

			internal FakeRange (bool throwOnWrite)
			{
				_throwOnWrite = throwOnWrite;
			}

			public string Text
			{
				get => _text;
				set
				{
					if (_throwOnWrite) {
						throw new InvalidOperationException ("write failed");
					}
					_text = value ?? string.Empty;
				}
			}
		}
	}
}
