using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DocxGrider
{
	public class DocxGrider : IDisposable
	{
		private WordprocessingDocument _document;
		private MemoryStream _memoryStream;

		/// <summary>
		/// Load document from stream.
		/// </summary>
		/// <param name="stream">Stream.</param>
		public DocxGrider(Stream stream)
		{
			if (stream == null)
			{
				new ArgumentNullException(nameof(stream));
			}

			LoadDocument(stream);
		}

		/// <summary>
		/// Load document from file.
		/// </summary>
		/// <param name="filename">File path.</param>
		public DocxGrider(string filename)
		{
			if (string.IsNullOrEmpty(filename))
			{
				new ArgumentException($"{nameof(filename)} is empty.");
			}

			using (var fileStream = new FileStream(filename, FileMode.Open, FileAccess.Read))
			{
				LoadDocument(fileStream);
			}
		}

		private void LoadDocument(Stream stream)
		{
			_memoryStream = new MemoryStream();
			stream.CopyTo(_memoryStream);
			_memoryStream.Position = 0;
			_document = WordprocessingDocument.Open(_memoryStream, true);
		}

		public void Dispose()
		{
			_document.Dispose();
			_memoryStream.Dispose();
		}

		/// <summary>
		/// Returns DocumentFormat.OpenXml document.
		/// </summary>
		/// <returns>DocumentFormat.OpenXml document.</returns>
		public WordprocessingDocument GetXmlDocument()
		{
			return _document;
		}

		public void SaveToStream(Stream stream)
		{
			if (stream == null)
			{
				new ArgumentNullException(nameof(stream));
			}

			_document.Save();
			_memoryStream.Position = 0;
			_memoryStream.CopyTo(stream);
		}

		public void SaveToFile(string filename)
		{
			if (string.IsNullOrEmpty(filename))
			{
				new ArgumentException($"{nameof(filename)} is empty.");
			}

			using (var fileStream = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.Write))
			{
				SaveToStream(fileStream);
			}
		}

		/// <summary>
		/// Replaces the first occurrence of the text.
		/// </summary>
		/// <param name="oldValue">Old value.</param>
		/// <param name="newValue">New value.</param>
		public void ReplaceText(string oldValue, string newValue)
		{
			ReplaceText(_document.MainDocumentPart.Document.Body, oldValue, newValue);
		}

		/// <summary>
		/// Replaces the first occurrence of the text, starts to search from the <paramref name="element"/> top element.
		/// </summary>
		/// <param name="element">Element to search inside from.</param>
		/// <param name="oldValue">Old value.</param>
		/// <param name="newValue">New value.</param>
		public void ReplaceText(OpenXmlElement element, string oldValue, string newValue)
		{
			if (element == null)
			{
				new ArgumentNullException(nameof(element));
			}
			if (string.IsNullOrEmpty(oldValue))
			{
				new ArgumentException($"{nameof(oldValue)} is empty.");
			}

			var sb = new StringBuilder();
			ReplaceTextChildren(new OpenXmlElementList(element), oldValue, newValue ?? string.Empty, sb);
		}

		private void ReplaceTextChildren(OpenXmlElementList children, string oldValue, string newValue, StringBuilder stringBuilder)
		{
			foreach (var child in children)
			{
				var textElements = new List<Text>();
				foreach (var subChild in child.ChildElements)
				{
					if (subChild is Run run)
					{
						foreach (var subSubChild in run.ChildElements)
						{
							if (subSubChild is Text textElement)
							{
								textElements.Add(textElement);
							}
						}
					}
				}

				textElements.ForEach(r => stringBuilder.Append(r.Text));
				// text from all Text elements together
				var text = stringBuilder.ToString();
				// position of oldValue in this text
				var targetTextsPos = text.IndexOf(oldValue);

				if (targetTextsPos != -1)
				{
					// Replace text in Run(s).
					// Text may be split into multiple Run\Text elements.
					// If multiple then put all new text in the first Run\Text element, and remove old text from other Run\Text elements.

					// example:
					// index          =  0         10        20
					// index          =  01234567890123456789012 = 23 chars
					// text           = "this is example of text"
					// oldValue       =         "example"
					// Text1          = "this is exa"
					// Text2          =            "mple of"
					// Text3          =                   " text"
					// targetTextsPos =          18
					int oldValuePos = 0;
					int textsPos = 0;

					// 1) Find position of oldValue in the chunks of the texts (textElements).
					// 2) Put newValue in the first found chunk.
					// 3) If there are other chunks of oldValue then replace their chars with string.Empty.
					for (int i = 0; i < textElements.Count; i++)
					{
						if (oldValuePos >= oldValue.Length)
						{
							break;
						}

						var textElement = textElements[i];
						var thisText = textElement.Text;

						// targetTextsPos is not reached with this textElement
						if (oldValuePos == 0 && textsPos + thisText.Length <= targetTextsPos)
						{
							textsPos += thisText.Length;
							continue;
						}

						if (oldValuePos == 0)
						{
							// here: first chunk with oldValue

							if (thisText.Length - (targetTextsPos - textsPos) >= oldValue.Length)
							{
								// the whole oldValue is in this textElement

								textElement.Text = thisText.Replace(oldValue, newValue);
								break;
							}
							else
							{
								// only the beginning of oldValue is in this textElement

								textElement.Text = thisText.Substring(0, targetTextsPos - textsPos) + newValue;
								textsPos += thisText.Length;
								oldValuePos += textsPos - targetTextsPos;
								continue;
							}
						}
						else
						{
							// here: not first chunk with oldValue

							if (thisText.Length >= (oldValue.Length - oldValuePos))
							{
								// the whole rest of the oldValue is in this textElement

								textElement.Text = thisText.Substring(oldValue.Length - oldValuePos);
								break;
							}
							else
							{
								// only part of the rest of oldValue is in this textElement

								oldValuePos += thisText.Length;
								textElement.Text = string.Empty;
								continue;
							}
						}
					}
				}

				stringBuilder.Clear();
				ReplaceTextChildren(child.ChildElements, oldValue, newValue, stringBuilder);
			}
		}

		/// <summary>
		/// Returns all tables that are first-level from the <paramref name="element"/>.
		/// </summary>
		/// <param name="element">First-level element to search from.</param>
		/// <returns>First-level tables.</returns>
		public List<Table> GetParentTables(OpenXmlElement element)
		{
			if (element == null)
			{
				new ArgumentNullException(nameof(element));
			}

			return GetParentTablesInner(element.ChildElements, new List<Table>());
		}

		private List<Table> GetParentTablesInner(OpenXmlElementList elements, List<Table> tables)
		{
			foreach (var element in elements)
			{
				if (element is Table table)
				{
					tables.Add(table);
				}
				else
				{
					GetParentTablesInner(element.ChildElements, tables);
				}
			}

			return tables;
		}

		/// <summary>
		/// Inserts copy of another row.
		/// </summary>
		/// <param name="table">Table.</param>
		/// <param name="sourceRowIndex">Source row index.</param>
		/// <param name="targetRowIndex">Row index before which the copy will be inserted.</param>
		public void InsertRowCopyBefore(Table table, int sourceRowIndex, int targetRowIndex)
		{
			var rows = table.ChildElements.OfType<TableRow>().ToList();
			var sourceRow = rows[sourceRowIndex];
			var newRow = (TableRow)sourceRow.Clone();
			rows[targetRowIndex].InsertBeforeSelf(newRow);
		}

		/// <summary>
		/// Returns rows of the table.
		/// </summary>
		/// <param name="table">Table.</param>
		/// <returns>Rows.</returns>
		public List<TableRow> GetTableRows(Table table)
		{
			var rows = table.ChildElements.OfType<TableRow>().ToList();
			return rows;
		}

		/// <summary>
		/// Replaces first occurrence of text, beginning to search from the <paramref name="element"/> top element.
		/// WARNING: if <paramref name="oldValue"/> text is inside another text withing the found Text element then the whole Text element will be replaced.
		/// </summary>
		/// <param name="element">Element to search inside from.</param>
		/// <param name="oldValue">Text to search.</param>
		/// <param name="image">Image to replace text with.</param>
		/// <param name="imageType">Image type, for example <see cref="ImagePartType.Jpeg"/>.</param>
		public void ReplaceText(OpenXmlElement element, string oldValue, byte[] image, PartTypeInfo imageType)
		{
			if (element == null)
			{
				new ArgumentNullException(nameof(element));
			}
			if (string.IsNullOrEmpty(oldValue))
			{
				new ArgumentException($"{nameof(oldValue)} is empty.");
			}
			if (image == null)
			{
				new ArgumentNullException(nameof(image));
			}
			if (imageType == null)
			{
				new ArgumentNullException(nameof(imageType));
			}

			// how to insert image:
			// https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-picture-into-a-word-processing-document

			throw new NotImplementedException();
		}
	}
}
