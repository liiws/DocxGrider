using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;

namespace DocxGrider.Tests
{
	[TestClass]
	public class TextReplaceTests
	{
		[TestMethod]
		public void AllTextInsideOneTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run = paragraph.AppendChild(new Run());
			var text = new Text("Some text element");
			run.AppendChild(text);

			// A

			dxg.ReplaceText(" text ", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts = runs[0].Elements<Text>().ToList();
			var textElement = texts[0];

			Assert.AreEqual("Some-new-element", textElement.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(1, runs.Count);
			Assert.AreEqual(1, texts.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		[TestMethod]
		public void TextInTheEndOfOneTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run = paragraph.AppendChild(new Run());
			var text = new Text("Some text element");
			run.AppendChild(text);

			// A

			dxg.ReplaceText("element", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts = runs[0].Elements<Text>().ToList();
			var textElement = texts[0];

			Assert.AreEqual("Some text -new-", textElement.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(1, runs.Count);
			Assert.AreEqual(1, texts.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		[TestMethod]
		public void TextIsTheWholeTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run = paragraph.AppendChild(new Run());
			var text = new Text("Some text element");
			run.AppendChild(text);

			// A

			dxg.ReplaceText("Some text element", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts = runs[0].Elements<Text>().ToList();
			var textElement = texts[0];

			Assert.AreEqual("-new-", textElement.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(1, runs.Count);
			Assert.AreEqual(1, texts.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		[TestMethod]
		public void TextInTheWholeFirstAndTheWholeNextTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run1 = paragraph.AppendChild(new Run());
			var text1 = new Text("Some ");
			run1.AppendChild(text1);
			var run2 = paragraph.AppendChild(new Run());
			var text2 = new Text("text element");
			run2.AppendChild(text2);

			// A

			dxg.ReplaceText("Some text element", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts1 = runs[0].Elements<Text>().ToList();
			var textElement1 = texts1[0];
			var texts2 = runs[1].Elements<Text>().ToList();
			var textElement2 = texts2[0];

			Assert.AreEqual("-new-", textElement1.Text);
			Assert.AreEqual("", textElement2.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(2, runs.Count);
			Assert.AreEqual(1, texts1.Count);
			Assert.AreEqual(1, texts2.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		[TestMethod]
		public void TextInTheWholeFirstAndTheWholeNextAndTheWholeNextTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run1 = paragraph.AppendChild(new Run());
			var text1 = new Text("Some ");
			run1.AppendChild(text1);
			var run2 = paragraph.AppendChild(new Run());
			var text2 = new Text("text ");
			run2.AppendChild(text2);
			var run3 = paragraph.AppendChild(new Run());
			var text3 = new Text("element");
			run3.AppendChild(text3);

			// A

			dxg.ReplaceText("Some text element", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts1 = runs[0].Elements<Text>().ToList();
			var textElement1 = texts1[0];
			var texts2 = runs[1].Elements<Text>().ToList();
			var textElement2 = texts2[0];
			var texts3 = runs[2].Elements<Text>().ToList();
			var textElement3 = texts3[0];

			Assert.AreEqual("-new-", textElement1.Text);
			Assert.AreEqual("", textElement2.Text);
			Assert.AreEqual("", textElement3.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(3, runs.Count);
			Assert.AreEqual(1, texts1.Count);
			Assert.AreEqual(1, texts2.Count);
			Assert.AreEqual(1, texts3.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		[TestMethod]
		public void TextInTheWholeFirstAndStartOfNextTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run1 = paragraph.AppendChild(new Run());
			var text1 = new Text("Some ");
			run1.AppendChild(text1);
			var run2 = paragraph.AppendChild(new Run());
			var text2 = new Text("text element");
			run2.AppendChild(text2);

			// A

			dxg.ReplaceText("Some text", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts1 = runs[0].Elements<Text>().ToList();
			var textElement1 = texts1[0];
			var texts2 = runs[1].Elements<Text>().ToList();
			var textElement2 = texts2[0];

			Assert.AreEqual("-new-", textElement1.Text);
			Assert.AreEqual(" element", textElement2.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(2, runs.Count);
			Assert.AreEqual(1, texts1.Count);
			Assert.AreEqual(1, texts2.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		[TestMethod]
		public void TextInTheEndOfFirstAndStartOfNextTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run1 = paragraph.AppendChild(new Run());
			var text1 = new Text("Some ");
			run1.AppendChild(text1);
			var run2 = paragraph.AppendChild(new Run());
			var text2 = new Text("text element");
			run2.AppendChild(text2);

			// A

			dxg.ReplaceText("me te", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts1 = runs[0].Elements<Text>().ToList();
			var textElement1 = texts1[0];
			var texts2 = runs[1].Elements<Text>().ToList();
			var textElement2 = texts2[0];

			Assert.AreEqual("So-new-", textElement1.Text);
			Assert.AreEqual("xt element", textElement2.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(2, runs.Count);
			Assert.AreEqual(1, texts1.Count);
			Assert.AreEqual(1, texts2.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		[TestMethod]
		public void TextInTheEndOfFirstAndWholeNextAndStartOfNextTextElement()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var paragraph = body.AppendChild(new Paragraph());
			var run1 = paragraph.AppendChild(new Run());
			var text1 = new Text("Some ");
			run1.AppendChild(text1);
			var run2 = paragraph.AppendChild(new Run());
			var text2 = new Text("text ");
			run2.AppendChild(text2);
			var run3 = paragraph.AppendChild(new Run());
			var text3 = new Text("element");
			run3.AppendChild(text3);

			// A

			dxg.ReplaceText("me text el", "-new-");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var newBody = resultDocument.MainDocumentPart.Document.Body;
			var paragraphs = newBody.Elements<Paragraph>().ToList();
			var runs = paragraphs[0].Elements<Run>().ToList();
			var texts1 = runs[0].Elements<Text>().ToList();
			var textElement1 = texts1[0];
			var texts2 = runs[1].Elements<Text>().ToList();
			var textElement2 = texts2[0];
			var texts3 = runs[2].Elements<Text>().ToList();
			var textElement3 = texts3[0];

			Assert.AreEqual("So-new-", textElement1.Text);
			Assert.AreEqual("", textElement2.Text);
			Assert.AreEqual("ement", textElement3.Text);
			Assert.AreEqual(1, paragraphs.Count);
			Assert.AreEqual(3, runs.Count);
			Assert.AreEqual(1, texts1.Count);
			Assert.AreEqual(1, texts2.Count);
			Assert.AreEqual(1, texts3.Count);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}

		private DocxGrider TestStart(out WordprocessingDocument srcDocument)
		{
			var dxg = CreateEmptyDocument();
			srcDocument = dxg.GetXmlDocument();
			return dxg;
		}

		private void TestEnd(DocxGrider dxg, WordprocessingDocument resultDocument, MemoryStream resultMemoryStream)
		{
			dxg.Dispose();
			resultDocument.Dispose();
			resultMemoryStream.Dispose();
		}

		private WordprocessingDocument TestGetResult(DocxGrider dxg, out MemoryStream resultMemoryStream)
		{
			resultMemoryStream = new MemoryStream();
			var resultDocument = ExportDocument(dxg, resultMemoryStream);
			return resultDocument;
		}

		private DocxGrider CreateEmptyDocument()
		{
			using (var memoryStream = new MemoryStream())
			{
				using (var document = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document))
				{
					var mainDocumentPart = document.AddMainDocumentPart();
					mainDocumentPart.Document = new Document();
					var body = document.MainDocumentPart.Document.AppendChild(new Body());
					document.Save();

					memoryStream.Position = 0;
					var dxg = new DocxGrider(memoryStream);

					return dxg;
				}
			}
		}

		private WordprocessingDocument ExportDocument(DocxGrider dxg, MemoryStream memoryStream)
		{
			dxg.SaveToStream(memoryStream);
			memoryStream.Position = 0;
			var document = WordprocessingDocument.Open(memoryStream, true);
			return document;
		}
	}
}
