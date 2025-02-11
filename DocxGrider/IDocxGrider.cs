using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;

namespace DocxGrider
{
	/// <summary>
	/// Works with .docx document.
	/// </summary>
	public interface IDocxGrider : IDisposable
	{
		/// <summary>
		/// Returns DocumentFormat.OpenXml document.
		/// </summary>
		/// <returns>DocumentFormat.OpenXml document.</returns>
		WordprocessingDocument GetXmlDocument();

		/// <summary>
		/// Saves document to the stream.
		/// </summary>
		/// <param name="stream">Stream to save to.</param>
		/// <param name="password">Password or null.</param>
		void SaveToStream(Stream stream, string password = null);

		/// <summary>
		/// Saves document to the file.
		/// </summary>
		/// <param name="filename">Filename to save to.</param>
		/// <param name="password">Password or null.</param>
		void SaveToFile(string filename, string password = null);

		/// <summary>
		/// Replaces the text.
		/// </summary>
		/// <param name="oldValue">Old value.</param>
		/// <param name="newValue">New value.</param>
		void ReplaceText(string oldValue, string newValue);

		/// <summary>
		/// Replaces the text, starts to search from the <paramref name="element"/> top element.
		/// </summary>
		/// <param name="element">Element to search inside from.</param>
		/// <param name="oldValue">Old value.</param>
		/// <param name="newValue">New value.</param>
		void ReplaceText(OpenXmlElement element, string oldValue, string newValue);

		/// <summary>
		/// Returns all tables that are first-level from the body.
		/// </summary>
		/// <param name="element">First-level element to search from.</param>
		/// <returns>First-level tables.</returns>
		List<Table> GetParentTables();

		/// <summary>
		/// Returns all tables that are first-level from the <paramref name="element"/>.
		/// </summary>
		/// <param name="element">First-level element to search from.</param>
		/// <returns>First-level tables.</returns>
		List<Table> GetParentTables(OpenXmlElement element);

		/// <summary>
		/// Inserts copy of another row.
		/// </summary>
		/// <param name="table">Table.</param>
		/// <param name="sourceRowIndex">Source row index.</param>
		/// <param name="targetRowIndex">Row index before which the copy will be inserted.</param>
		TableRow InsertRowCopyBefore(Table table, int sourceRowIndex, int targetRowIndex);

		/// <summary>
		/// Inserts copy of another row.
		/// </summary>
		/// <param name="table">Table.</param>
		/// <param name="sourceRowIndex">Source row index.</param>
		/// <param name="targetRowIndex">Row index after which the copy will be inserted.</param>
		TableRow InsertRowCopyAfter(Table table, int sourceRowIndex, int targetRowIndex);

		/// <summary>
		/// Removes specified table row.
		/// </summary>
		/// <param name="table">Table.</param>
		/// <param name="iowIndex">Row index to remove.</param>
		void RemoveTableRow(Table table, int rowIndex);

		/// <summary>
		/// Returns rows of the table.
		/// </summary>
		/// <param name="table">Table.</param>
		/// <returns>Rows.</returns>
		List<TableRow> GetTableRows(Table table);

		/// <summary>
		/// Finds first element that has specified alternative text.
		/// </summary>
		/// <param name="text">Text.</param>
		/// <param name="element">Element to start from, or document root if not specified.</param>
		/// <returns>Found element or null.</returns>
		OpenXmlElement FindElementWithAlternativeText(string text, OpenXmlElement element = null);

		/// <summary>
		/// Returns all elements of the specified type.
		/// </summary>
		/// <typeparam name="T">Type.</typeparam>
		/// <param name="element">Element to start from, or document root if not specified.</param>
		/// <returns>Elements.</returns>
		List<T> GetAllElements<T>(OpenXmlElement element = null) where T : OpenXmlElement;

		/// <summary>
		/// Removes part of the document (only <see cref="Paragraph"/>, <see cref="Run"/>, <see cref="Table"/>) between page breaks.
		/// </summary>
		/// <param name="sectionIndex">0 to remove from beginning to the first page break, 1 from first to second page break, and so on.</param>
		/// <returns>True if document part was removed, False otherwise.</returns>
		bool RemovePageBreakPart(int sectionIndex);
	}
}
