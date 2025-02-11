# DocxGrider

DocxGrider is a very simple library for working with templates in .docx format.

## Features

* Replace text including text box ([NPOI can't yet](https://github.com/nissl-lab/npoi/issues/1478))
* Insert copy of the row in the table at specified row index
* Delete table row
* Load from file or stream
* Save to file or stream
* Protect document with password (MSO 2007 legacy implementation)
* Very fast and doesn't require Office installed (uses Open XML SDK)

Not much, but enough for many templating purposes.

In any case it's possible to get OpenXML document with `GetXmlDocument()` and look at some [examples](https://github.com/OfficeDev/open-xml-docs/tree/main/samples/word) how to work with it.

## How to use

Download sources or install [nuget package](https://www.nuget.org/packages/DocxGrider).

```cs
using (var fs = new FileStream("delivery_template.docx", FileMode.Open, FileAccess.Read))
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