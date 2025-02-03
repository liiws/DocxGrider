using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace DocxGrider
{
	public class DocxGrider : IDocxGrider
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

		/// <inheritdoc/>
		public WordprocessingDocument GetXmlDocument()
		{
			return _document;
		}

		/// <inheritdoc/>
		public void SaveToStream(Stream stream, string password = null)
		{
			if (stream == null)
			{
				new ArgumentNullException(nameof(stream));
			}

			if (string.IsNullOrEmpty(password))
			{
				RemoveProtection();
			}
			else
			{
				AddProtection(password);
			}

			_document.Save();
			_memoryStream.Position = 0;
			_memoryStream.CopyTo(stream);
		}

		/// <inheritdoc/>
		public void SaveToFile(string filename, string password = null)
		{
			if (string.IsNullOrEmpty(filename))
			{
				new ArgumentException($"{nameof(filename)} is empty.");
			}

			using (var fileStream = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.Write))
			{
				SaveToStream(fileStream, password);
			}
		}

		// The initial code array contains the initial values for the key’s high-order word. The initial value depends on the length of the password, as follows:
		private static readonly ushort[] _initialCodeArray =
		{
			0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C,
			0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139,
			0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3
		};

		private static readonly ushort[] _encryptionMatrix =
		{
			// bit 0     1       2       3       4       5       6
			0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, // last-14
			0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF, // last-13
			0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, // last-12
			0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, // last-11
			0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, // last-10
			0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, // last-9
			0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, // last-8
			0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, // last-7
			0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, // last-6
			0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, // last-5
			0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, // last-4
			0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, // last-3
			0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, // last-2
			0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, // last-1
			0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4, // last
		};

		private void AddProtection(string password)
		{
			// Document protection
			// ECMA-376 Office Open XML file formats
			// https://go.microsoft.com/fwlink/?LinkId=200054

			// Write Protection Method
			// [MS-OFFCRYPTO]: Office Document Cryptography Structure
			// https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/3c34d72a-1a61-4b52-a893-196f9157f083

			// Password Hashing:
			// ECMA-376 Part 4 Transitional Migration Features

			// Document Editing Restrictions for Word:
			// MS-OE376: Office Implementation Information for ECMA-376 Standards Support
			// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/db9b9b72-b10b-4e7e-844c-09f88c972219

			const int maxPasswordLength = 15;
			const int spinCount = 50000;

			var salt = new byte[16];
			new RNGCryptoServiceProvider().GetNonZeroBytes(salt);

			// Truncate the password to 15 characters.
			if (password.Length > maxPasswordLength)
			{
				password = password.Substring(0, maxPasswordLength);
			}

			// Construct a new NULL-terminated string consisting of single-byte characters:
			byte[] passwordSingleBytesString = new byte[password.Length];
			// Get the single-byte values by iterating through the Unicode characters of the truncated password.
			// For each character, if the low byte is not equal to 0, take it. Otherwise, take the high byte.
			for (int i = 0; i < password.Length; i++)
			{
				passwordSingleBytesString[i] = (byte)password[i];
				if (passwordSingleBytesString[i] == 0)
				{
					passwordSingleBytesString[i] = (byte)(password[i] >> 8);
				}
			}

			// From now on, the single-byte character string is used.

			// Compute the high-order word of the new key:
			// Initialize from the initial code array (see below), depending on the password’s length.
			ushort highOrderWord = _initialCodeArray[password.Length - 1];
			// For each character in the password:
			for (int iChar = 0; iChar < passwordSingleBytesString.Length; iChar++)
			{
				// For every bit in the character, starting with the least significant and progressing to (but excluding) the most significant,
				// if the bit is set, XOR the key’s high-order word with the corresponding word from the encryption matrix
				for (int iBit = 0; iBit < 7; iBit++)
				{
					if (((passwordSingleBytesString[iChar] >> iBit) & 1) == 1)
					{
						var row = maxPasswordLength - passwordSingleBytesString.Length + iChar;
						var col = iBit;
						highOrderWord = (ushort)(highOrderWord ^ _encryptionMatrix[row * 7 + col]);
					}
				}
			}

			// Compute the low-order word of the new key:
			// Initialize with 0
			ushort lowOrderWord = 0;
			for (int i = passwordSingleBytesString.Length - 1; i >= 0; i--)
			{
				// For each character in the password, going backwards,
				// low-order word = (((low-order word SHR 14) AND 0x0001) OR (low-order word SHL 1) AND 0x7FFF)) XOR character
				lowOrderWord = (ushort)((((lowOrderWord >> 14) & 0x0001) | (lowOrderWord << 1) & 0x7FFF) ^ passwordSingleBytesString[i]);
			}
			// Lastly,
			// low-order word = (((low-order word SHR 14) AND 0x0001) OR (low-order word SHL 1) AND 0x7FFF)) XOR password length XOR 0xCE4B.
			lowOrderWord = (ushort)(((((lowOrderWord >> 14) & 0x0001) | (lowOrderWord << 1) & 0x7FFF)) ^ passwordSingleBytesString.Length ^ 0xCE4B);

			int passwordInitHash = (highOrderWord << 16) | lowOrderWord;

			// Second, the byte order of the result shall be reversed
			// example: 0x64CEED7E becomes 7EEDCE64.

			// In this third stage, the reversed byte order legacy hash from the second stage shall be
			// converted to Unicode hex string representation [Example: If the single byte string 7EEDCE64 is
			// converted to Unicode hex string it will be represented in memory as the following byte stream: 37 00
			// 45 00 45 00 44 00 43 00 45 00 36 00 34 00. end example], and that value shall be hashed as defined
			// by the attribute values.
			// This note applies to the following products: 2007, 2007 SP1, 2007 SP2.

			byte[] passwordReversedHash = new byte[4]
			{
				(byte)(passwordInitHash & 0xFF),
				(byte)((passwordInitHash >> 8) & 0xFF),
				(byte)((passwordInitHash >> 16) & 0xFF),
				(byte)((passwordInitHash >> 24) & 0xFF),
			};

			byte[] passwordHash = Encoding.Unicode.GetBytes(string.Join("", passwordReversedHash.Select(r => r.ToString("X2"))));

			// H0 = H(salt + password)
			HashAlgorithm sha1 = new SHA1Managed();
			byte[] hash = sha1.ComputeHash(salt.Concat(passwordHash).ToArray());

			// Hn = H(Hn-1 + iterator)
			// where iterator is initially set to 0 and is incremented monotonically on each iteration until SpinCount
			// iterations have been performed. The value of iterator on the last iteration MUST be one less than
			// SpinCount. The final hash is then Hfinal = HSpinCount-1.
			for (int i = 0; i < spinCount; i++)
			{
				byte[] bytesCount = new byte[] { (byte)i, (byte)(i >> 8), (byte)(i >> 16), (byte)(i >> 24) };
				hash = sha1.ComputeHash(hash.Concat(bytesCount).ToArray());
			}

			var settingsPart = _document.MainDocumentPart.DocumentSettingsPart ?? _document.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
			var documentProtection = settingsPart.Settings?.GetFirstChild<DocumentProtection>() ?? new DocumentProtection();
			documentProtection.Edit = DocumentProtectionValues.ReadOnly;
			documentProtection.Enforcement = new OnOffValue(true);
			documentProtection.CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash;
			documentProtection.CryptographicProviderType = CryptProviderValues.RsaFull;
			documentProtection.CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny;
			documentProtection.CryptographicAlgorithmSid = 4; // SHA1
			documentProtection.CryptographicSpinCount = (uint)spinCount;
			documentProtection.Hash = Convert.ToBase64String(hash);
			documentProtection.Salt = Convert.ToBase64String(salt);
			_document.MainDocumentPart.DocumentSettingsPart.Settings.AppendChild(documentProtection);
		}

		private void RemoveProtection()
		{
			if (_document.MainDocumentPart?.DocumentSettingsPart?.Settings != null)
			{
				_document.MainDocumentPart.DocumentSettingsPart.Settings.RemoveAllChildren<DocumentProtection>();
			}
		}

		/// <inheritdoc/>
		public void ReplaceText(string oldValue, string newValue)
		{
			ReplaceText(_document.MainDocumentPart.Document.Body, oldValue, newValue);
		}

		/// <inheritdoc/>
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

		/// <inheritdoc/>
		public List<Table> GetParentTables()
		{
			return GetParentTablesInner(new OpenXmlElementList(_document.MainDocumentPart.Document.Body), new List<Table>());
		}

		/// <inheritdoc/>
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

		/// <inheritdoc/>
		public TableRow InsertRowCopyBefore(Table table, int sourceRowIndex, int targetRowIndex)
		{
			var rows = table.ChildElements.OfType<TableRow>().ToList();
			var sourceRow = rows[sourceRowIndex];
			var newRow = (TableRow)sourceRow.Clone();
			rows[targetRowIndex].InsertBeforeSelf(newRow);
			return newRow;
		}

		/// <inheritdoc/>
		public TableRow InsertRowCopyAfter(Table table, int sourceRowIndex, int targetRowIndex)
		{
			var rows = table.ChildElements.OfType<TableRow>().ToList();
			var sourceRow = rows[sourceRowIndex];
			var newRow = (TableRow)sourceRow.Clone();
			rows[targetRowIndex].InsertAfterSelf(newRow);
			return newRow;
		}

		/// <inheritdoc/>
		public void RemoveTableRow(Table table, int rowIndex)
		{
			var rows = table.ChildElements.OfType<TableRow>().ToList();
			rows[rowIndex].Remove();
		}

		/// <inheritdoc/>
		public List<TableRow> GetTableRows(Table table)
		{
			var rows = table.ChildElements.OfType<TableRow>().ToList();
			return rows;
		}

		/// <summary>
		/// Replaces the text to the image, beginning to search from the <paramref name="element"/> top element.
		/// WARNING: if <paramref name="oldValue"/> text is inside another text withing the found Text element then the whole Text element will be replaced.
		/// </summary>
		/// <param name="element">Element to search inside from.</param>
		/// <param name="oldValue">Text to search.</param>
		/// <param name="image">Image to replace text with.</param>
		/// <param name="imageType">Image type, for example <see cref="ImagePartType.Jpeg"/>.</param>
		private void ReplaceText(OpenXmlElement element, string oldValue, byte[] image, PartTypeInfo imageType)
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
