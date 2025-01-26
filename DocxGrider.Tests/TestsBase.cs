using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocxGrider.Tests
{
	public class TestsBase
	{
		protected DocxGrider TestStart(out WordprocessingDocument srcDocument)
		{
			var dxg = CreateEmptyDocument();
			srcDocument = dxg.GetXmlDocument();
			return dxg;
		}

		protected void TestEnd(DocxGrider dxg, WordprocessingDocument resultDocument, MemoryStream resultMemoryStream)
		{
			dxg.Dispose();
			resultDocument.Dispose();
			resultMemoryStream.Dispose();
		}

		protected WordprocessingDocument TestGetResult(DocxGrider dxg, out MemoryStream resultMemoryStream)
		{
			resultMemoryStream = new MemoryStream();
			var resultDocument = ExportDocument(dxg, resultMemoryStream);
			return resultDocument;
		}

		protected DocxGrider CreateEmptyDocument()
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

		protected WordprocessingDocument ExportDocument(DocxGrider dxg, MemoryStream memoryStream)
		{
			dxg.SaveToStream(memoryStream);
			memoryStream.Position = 0;
			var document = WordprocessingDocument.Open(memoryStream, true);
			return document;
		}
	}
}
