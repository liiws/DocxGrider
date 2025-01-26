using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace DocxGrider.Tests
{
	[TestClass]
	public class TableTests : TestsBase
	{
		[TestMethod]
		public void CopyRow()
		{
			// A

			var dxg = TestStart(out var srcDocument);

			var body = srcDocument.MainDocumentPart.Document.Body;
			var table = body.AppendChild(new Table());
			var row = table.AppendChild(new TableRow());
			row.Append(new TableCell(new Paragraph(new Run(new Text("{Id}")))));
			row.Append(new TableCell(new Paragraph(new Run(new Text("{Name}")))));

			// A

			var dxgTable = dxg.GetParentTables(body)[0];
			var newRow = dxg.InsertRowCopyBefore(dxgTable, 0, 0);
			dxg.ReplaceText(newRow, "{Id}", "1");
			dxg.ReplaceText(newRow, "{Name}", "Name1");
			var dxgRows = dxg.GetTableRows(dxgTable);
			dxg.ReplaceText(dxgRows[1], "{Id}", "2");
			dxg.ReplaceText(dxgRows[1], "{Name}", "Name2");

			// A

			var resultDocument = TestGetResult(dxg, out var resultMemoryStream);
			var resultBody = resultDocument.MainDocumentPart.Document.Body;
			var tables = body.Elements<Table>().ToList();
			var rows = table.Elements<TableRow>().ToList();
			var row1Cells = rows[0].Elements<TableCell>().ToList();
			var row2Cells = rows[1].Elements<TableCell>().ToList();
			var row1Cell1Paragraphs = row1Cells[0].Elements<Paragraph>().ToList();
			var row1Cell2Paragraphs = row1Cells[1].Elements<Paragraph>().ToList();
			var row2Cell1Paragraphs = row2Cells[0].Elements<Paragraph>().ToList();
			var row2Cell2Paragraphs = row2Cells[1].Elements<Paragraph>().ToList();
			var row1Cell1Paragraph1Runs = row1Cell1Paragraphs[0].Elements<Run>().ToList();
			var row1Cell2Paragraph1Runs = row1Cell2Paragraphs[0].Elements<Run>().ToList();
			var row2Cell1Paragraph1Runs = row2Cell1Paragraphs[0].Elements<Run>().ToList();
			var row2Cell2Paragraph1Runs = row2Cell2Paragraphs[0].Elements<Run>().ToList();
			var row1Cell1Paragraph1Run1Texts = row1Cell1Paragraph1Runs[0].Elements<Text>().ToList();
			var row1Cell2Paragraph1Run1Texts = row1Cell2Paragraph1Runs[0].Elements<Text>().ToList();
			var row2Cell1Paragraph1Run1Texts = row2Cell1Paragraph1Runs[0].Elements<Text>().ToList();
			var row2Cell2Paragraph1Run1Texts = row2Cell2Paragraph1Runs[0].Elements<Text>().ToList();
			var row1Cell1Paragraph1Run1Text1 = row1Cell1Paragraph1Run1Texts[0];
			var row1Cell2Paragraph1Run1Text1 = row1Cell2Paragraph1Run1Texts[0];
			var row2Cell1Paragraph1Run1Text1 = row2Cell1Paragraph1Run1Texts[0];
			var row2Cell2Paragraph1Run1Text1 = row2Cell2Paragraph1Run1Texts[0];

			Assert.AreEqual(1, tables.Count);
			Assert.AreEqual(2, rows.Count);
			Assert.AreEqual(2, row1Cells.Count);
			Assert.AreEqual(2, row2Cells.Count);
			Assert.AreEqual(1, row1Cell1Paragraphs.Count);
			Assert.AreEqual(1, row1Cell2Paragraphs.Count);
			Assert.AreEqual(1, row2Cell1Paragraphs.Count);
			Assert.AreEqual(1, row2Cell2Paragraphs.Count);
			Assert.AreEqual(1, row1Cell1Paragraph1Runs.Count);
			Assert.AreEqual(1, row1Cell2Paragraph1Runs.Count);
			Assert.AreEqual(1, row2Cell1Paragraph1Runs.Count);
			Assert.AreEqual(1, row2Cell2Paragraph1Runs.Count);
			Assert.AreEqual(1, row1Cell1Paragraph1Run1Texts.Count);
			Assert.AreEqual(1, row1Cell2Paragraph1Run1Texts.Count);
			Assert.AreEqual(1, row2Cell1Paragraph1Run1Texts.Count);
			Assert.AreEqual(1, row2Cell2Paragraph1Run1Texts.Count);
			Assert.AreEqual("1", row1Cell1Paragraph1Run1Text1.Text);
			Assert.AreEqual("Name1", row1Cell2Paragraph1Run1Text1.Text);
			Assert.AreEqual("2", row2Cell1Paragraph1Run1Text1.Text);
			Assert.AreEqual("Name2", row2Cell2Paragraph1Run1Text1.Text);

			TestEnd(dxg, resultDocument, resultMemoryStream);
		}
	}
}
