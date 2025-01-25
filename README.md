# DocxGrider

DocxGrider is a simple librabry for working with templates in .docx format.

## Features

* Replace text in body and text box
* Insert a copy of the row in table at specified row index
* Load from file or stream
* Save to file or stream

## How to use

```cs
using (var fs = new FileStream("textbox.docx", FileMode.Open, FileAccess.Read))
{
	using (var dxg = new DocxGrider(fs))
	{
		// replace text
		dxg.ReplaceText("{PhoneNumber}", "+1234567890");

		// duplicate row and replace text in it
		var table = dxg.GetParentTables()[0];		
		var newRow = dxg.InsertRowCopyBefore(table, 1, 1);

		// clone rows and replace text in them
		var itemsCount = 5;
		for (var i = 0; i < itemsCount - 1; i++)
		{
			dxg.InsertRowCopyBefore(table, 1, 1);
		}

		// fill clones rows
		var rows = GetTableRows(table);
		for (var i = 0; i < itemsCount; i++)
		{
			var row = rows[1 + i];
			dxg.ReplaceText(row, "{Index}", i.ToString());
		}

		// save to file
		var tempFile = Path.GetTempFileName() + ".docx";
		dxg.SaveToFile(tempFile);
	}
}
```