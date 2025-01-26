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
			var textElement = newBody.Elements<Paragraph>().First().Elements<Run>().First().Elements<Text>().First();

			Assert.AreEqual("Some-new-element", textElement.Text);

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
