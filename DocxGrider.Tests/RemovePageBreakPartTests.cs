using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace DocxGrider.Tests
{
	[TestClass]
	public class RemovePageBreakPartTests : TestsBase
	{
		[TestMethod]
		public void RemovePart0_SameParagraph()
		{
			// A

			var dxg = TestStart(out var srcDocument);
			{
				var body = srcDocument.MainDocumentPart.Document.Body;
				var paragraph1 = body.AppendChild(new Paragraph());
				var run1 = paragraph1.AppendChild(new Run());
				var text1 = new Text("Text before break");
				run1.AppendChild(text1);
				var run2 = paragraph1.AppendChild(new Run());
				var break2 = new Break();
				break2.Type = BreakValues.Page;
				run2.AppendChild(break2);
				var run3 = paragraph1.AppendChild(new Run());
				var text3 = new Text("Text after break");
				run3.AppendChild(text3);
			}

			// A

			dxg.RemovePageBreakPart(0);

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var resultBody = resultDocument.MainDocumentPart.Document.Body;
			{
				var paragraphs = resultBody.Elements<Paragraph>().ToList();
				var runs = paragraphs[0].Elements<Run>().ToList();
				var run2 = runs[0];
				var run3 = runs[1];
				var text3 = run3.Elements<Text>().First();

				Assert.AreEqual(1, paragraphs.Count);
				Assert.AreEqual(2, runs.Count);
				Assert.AreEqual(0, run2.ChildElements.Count);
				Assert.AreEqual(1, run3.ChildElements.Count);
				Assert.AreEqual("Text after break", text3.Text);

				TestEnd(dxg, resultDocument, resultMemoryStream);
			}
		}

		[TestMethod]
		public void RemovePart1_SameParagraph()
		{
			// A

			var dxg = TestStart(out var srcDocument);
			{
				var body = srcDocument.MainDocumentPart.Document.Body;
				var paragraph1 = body.AppendChild(new Paragraph());
				var run1 = paragraph1.AppendChild(new Run());
				var text1 = new Text("Text before break");
				run1.AppendChild(text1);
				var run2 = paragraph1.AppendChild(new Run());
				var break2 = new Break();
				break2.Type = BreakValues.Page;
				run2.AppendChild(break2);
				var run3 = paragraph1.AppendChild(new Run());
				var text3 = new Text("Text after break");
				run3.AppendChild(text3);
			}

			// A

			dxg.RemovePageBreakPart(1);

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var resultBody = resultDocument.MainDocumentPart.Document.Body;
			{
				var paragraphs = resultBody.Elements<Paragraph>().ToList();
				var runs = paragraphs[0].Elements<Run>().ToList();
				var run1 = runs[0];
				var run2 = runs[1];
				var text1 = run1.Elements<Text>().First();

				Assert.AreEqual(1, paragraphs.Count);
				Assert.AreEqual(2, runs.Count);
				Assert.AreEqual(1, run1.ChildElements.Count);
				Assert.AreEqual(0, run2.ChildElements.Count);
				Assert.AreEqual("Text before break", text1.Text);

				TestEnd(dxg, resultDocument, resultMemoryStream);
			}
		}
	}
}
