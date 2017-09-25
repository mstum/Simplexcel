# Simplexcel
This is a simple .xlsx generator library for .net 4.5 or .NET Standard 1.3 (or higher).  
  
It does not aim to implement the entirety of the Office Open XML Workbook format and all the small and big features Excel offers.  
Instead, it is meant as a way to handle common tasks that can't be handled by other workarounds (e.g., CSV Files or HTML Tables) and is fully supported under ASP.net (unlike, say, COM Interop which Microsoft explicitly doesn't support on a server).

# Features
* You can store numbers as numbers, so no more unwanted conversion to scientific notation on large numbers!
* You can store text that looks like a number as text, so no more truncation of leading zeroes because Excel thinks it's a number
* You can have multiple Worksheets
* You have basic formatting: Font Name/Size, Bold/Underline/Italic, Color, Border around a cell
* You can specify the size of cells
* Workbooks can be saved compressed or uncompressed (CPU Load vs. File Size)
* You can specify repeating rows and columns (from the top and left respectively), useful when printing.
* Fully supported in ASP.net and Windows Services
* Supports .net Core

# Not supported (yet?)
* Automatic Column Sizing
* Merging Cells
* Embedding any Media or anything OLE
* Creating Charts
* Reading Excel sheets

# Usage
## Install Nuget Package
The Library is available on NuGet and can be installed by searching for Simplexcel or using the NuGet command prompt:

    Install-Package simplexcel

For more information on NuGet see http://www.nuget.org

## Examples
### Creating a Workbook
The simple example of creating a workbook only requires you to create a few Worksheets, populate their Cells, and add them to a new Workbook.

```cs
// using Simplexcel;
var sheet = new Worksheet("Hello, world!");
sheet.Cells[0, 0] = "Hello,";
sheet.Cells["B1"] = "World!";

var workbook = new Workbook();
workbook.Add(sheet);
workbook.Save(@"d:\test.xlsx");
```

### Cell creation
Cells can be created implicitly from strings, int64's, double and decimal (and anything that's implicitly convertible to those, e.g. int32), DateTime, or explicitly through the Cell constructor.

```cs
Cell textCell = "fromString";
Cell intCell = 4; // will be formatted as number without decimal places "0"
Cell doubleCell = 123.456; // will be formatted with 2 decimal places, "0.00"
Cell decimalCell = 987.654321m; // will be formatted with 2 decimal places, "0.00"
decimalCell.Format = "0.000"; // override the cell format with an Excel Formatting string
Cell dateCell = DateTime.Now; // will be formatted with date and time portions
Cell cellFromStaticFactory = Cell.FromObject(myObj.SomeProp); // will try to guess the cell type based on the type of the object - note that this causes boxing

// Explicit constructor, specifying the type, value and format (Excel will display this as 88.75%)
Cell decimalWithFormat = new Cell(CellType.Number, 0.8875m, BuiltInCellFormat.PercentTwoDecimalPlaces);
```

### Cell formatting
After you've added a cell, you can set formatting like Bold or Font Size or Border on it.

```cs
// Addressing cells via zero-based row and column index...
sheet.Cells[0, 0] = "Hello,";
sheet.Cells[0, 0].Bold = true;
// ...or Excel-style references, B1 = [0,1]
sheet.Cells["B1"] = "World!";
sheet.Cells["B1"].Border = CellBorder.Bottom | CellBorder.Right;

Cell myCell = "Test Cell";
myCell.FontName = "Comic Sans MS";
myCell.FontSize = 18;
myCell.TextColor = Color.Violet;
myCell.Bold = true;
myCell.Italic = true;
myCell.Underline = true;
myCell.Border = CellBorder.All; // Left | Right | Top | Bottom
sheet.Cells[0, 2] = myCell;

// To change the width of a column, specify a value. Specifying NULL will use the default
sheet.ColumnWidths[2] = 30;
```
			
In Excel, this will create a beautiful sheet:  
![Cell formatting](/doc/formatting.png?raw=true "Cell formatting")

### Page Setup
A Worksheet has a PageSetup property that contains the orientation (portrait or landscape) and a setting to repeat a number of rows or columns (starting from the top and left respectively) on every page when printing.

```cs
var sheet = new Worksheet("Hello, world!");
sheet.PageSetup.PrintRepeatRows = 2; // How many rows (starting with the top one)
sheet.PageSetup.PrintRepeatColumns = 0; // How many columns (starting with the left one, 0 is default)
sheet.PageSetup.Orientation = Orientation.Landscape;


sheet.Cells["A1"] = "Title!";
sheet.Cells["A1"].Bold = true;
sheet.Cells["A2"] = "Subtitle!";
sheet.Cells["A2"].Bold = true;
sheet.Cells["A2"].TextColor = Color.Magenta;
for (int i = 0; i < 100; i++)
{
    sheet.Cells[i + 2, 0] = "Entry Number " + (i + 1);
}
```
			
This looks like this when printing:  
![Page Setup](/doc/repeatrows.png?raw=true "Page Setup")

### Hyperlinks
You can create Hyperlinks for a cell.

```cs
sheet.Cells["A1"] = "Click me now!";  
sheet.Cells["A1"].Hyperlink = "https://github.com/mstum/Simplexcel/";  
```

This will NOT automatically format it as a Hyperlink (blue/underlined) to give you freedom to format as desired.

### Misc.
You can specify if the .xlsx file should be compressed.

```cs
workbook.Save(@"d:\testCompressed.xlsx", true); // Smaller file, more CPU usage, Default
workbook.Save(@"d:\testUncompressed.xlsx", false); // Larger file, less CPU usage
```

You can also save to a Stream. The Stream must be writeable.
```
workbook.Save(HttpContext.Current.Response.OutputStream);
```

The workbook class allows you to add an Author and Title, which will show up in the Properties pane.

```cs
var workbook = new Workbook();
workbook.Title = "Hello, World!";
workbook.Author = "My Application Version 1.0.0.5";
```
![Title and Author](/doc/titleauthor.png?raw=true "Title and Author")

To help troubleshooting issues with generated xlsx files, they contain a file called simplexcel.xml which contains version information and the date the document was generated.

![simplexcel.xml](/doc/simplexcelxml.png?raw=true "simplexcel.xml")

```xml
<docInfo xmlns="http://stum.de/simplexcel/v1">
	<version major="2" minor="0" build="0" revision="0"></version>
	<created>2017-04-07T21:04:18.5220876Z</created>
</docInfo>
```

### Caveats and Exceptions
Here's a list of things to be aware of when working with Simplexcel

```cs
// Sheet names cannot be NULL or empty
new Worksheet(null); // ArgumentException
new Worksheet(""); // ArgumentException

// Sheet names cannot be longer than 31 chars
new Worksheet(new string('a', 32));  // ArgumentException

// Sheet names cannot contain these chars: \ / ? * [ ]
new Worksheet("Data for 09/22/2012"); // ArgumentException

// There are static properties on the Worksheet class to help you create valid names
int maxLength = Worksheet.MaxSheetNameLength;
char[] invalidChars = Worksheet.InvalidSheetNameChars;

// Within a workbook, sheet names must be unique
var wb = new Workbook();
var sheet1 = new Worksheet("Hello!");
var sheet2 = new Worksheet("Hello!");
wb.Add(sheet1);
wb.Add(sheet2); // ArgumentException

// Accessing a cell before it's been assigned returns NULL
sheet1.Cells["A1"].Bold = true; // NullReferenceException
```

## ActionResult for ASP.net MVC
If you want to generate Excel Sheets as part of an ASP.net MVC action, you can use this abstract class as an ActionResult. A derived class would simply override GenerateWorkbook. This assumes that the controller merely creates a model, and that the actual workbook generation happens after the action is executed and ActionResult.ExecuteResult is called. Alternatively, you can change it to take a Workbook in the constructor and create the Workbook in the MVC Controller Action.

```cs
public abstract class ExcelResultBase : ActionResult
{
    private readonly string _filename;

    protected ExcelResultBase(string filename)
    {
        _filename = filename;
    }

    protected abstract Workbook GenerateWorkbook();

    public override void ExecuteResult(ControllerContext context)
    {
        if (context == null)
        {
            throw new ArgumentNullException("context");
        }

        var workbook = GenerateWorkbook();
        if (workbook == null)
        {
            context.HttpContext.Response.StatusCode = (int)HttpStatusCode.NotFound;
            return;
        }

        context.HttpContext.Response.ContentType = "application/octet-stream";
        context.HttpContext.Response.StatusCode = (int)HttpStatusCode.OK;
        context.HttpContext.Response.AppendHeader("content-disposition", "attachment; filename=\"" + _filename + "\"");

        workbook.Save(context.HttpContext.Response.OutputStream);
    }
}
```

# Changelog
## 2.1.0 (In Development)
* **Functional Change:** Numbers with more than 11 digits are forced as Text by Default, because [of a limitation in Excel](https://support.microsoft.com/en-us/help/2643223/long-numbers-are-displayed-incorrectly-in-excel). To restore the previous functionality, you can set `Worksheet.LargeNumberHandlingMode` to `LargeNumberHandlingMode.None`. You can also use `Cell.IsLargeNumber` to check if a given number would be affected by this.
* **Functional Change:** `Worksheet.Populate`/`Worksheet.FromData` now also reads properties from base classes.
* `Worksheet.Populate`/`Worksheet.FromData` accept a new argument, `cacheTypeColumns` which defaults to false. If set to true, then Simplexcel will cache the Reflection-based lookup of object properties. This is useful for if you have a few types that you create sheets from a lot.
* You can add `[XlsxColumn]` to a Property so that `Worksheet.Populate`/`Worksheet.FromData` can set the column name and a given column order. *Caveat:* If you set `ColumnIndex` on some, but not all Properties, the properties without a `ColumnIndex` will be on the right of the last assigned column, even if that means gaps. I recommend that you either set `ColumnIndex` on all properties or none.
* You can add `[XlsxIgnoreColumn]` to a Property so that `Worksheet.Populate`/`Worksheet.FromData` ignores it.
* Added `Cell.HorizontalAlignment` and `Cell.VerticalAlignment` to allow setting the alignment of a cell (left/center/right/justify, top/middle/bottom/justify).
* Added XmlDoc to Nuget package, so you should get Intellisense with proper comments now.

## 2.0.5 (2017-09-23)
* Add support for manual page breaks. Call `Worksheet.InsertManualPageBreakAfterRow` or `Worksheet.InsertManualPageBreakAfterColumn` with either the zero-based index of the row/column after which to create the break, or with a cell address (e.g., B5) to create the break below or to the left of that cell.

## 2.0.4 (2017-09-17)
* Support for [freezing panes](https://support.office.com/en-us/article/Freeze-panes-to-lock-rows-and-columns-dab2ffc9-020d-4026-8121-67dd25f2508f). Right now, this is being kept simple: call either `Worksheet.FreezeTopRow` or `Worksheet.FreezeLeftColumn` to freeze either the first row (1) or the leftmost column (A).
* If a Stream is not seekable (e.g., HttpContext.Response.OutputStream), Simplexcel automatically creates a temporary MemoryStream as an intermediate.
* Add `Cell.FromObject` to make Cell creation easier by guessing the correct type.
* Support `DateTime` cells, thanks to @mguinness and PR #16.

## 2.0.3 (2017-09-08)
* Add `Worksheet.Populate<T>` method to fill a sheet with data. Caveats: Does not loot at inherited members, doesn't look at complex types.
* Also add static `Worksheet.FromData<T>` method to create and populate the sheet in one.

## 2.0.2 (2017-06-17)
* Add additional validation when saving to a Stream. The stream must be seekable (and of course writeable), otherwise an Exception is thrown.

## 2.0.1 (2017-05-18)
* Fix [Issue #12](https://github.com/mstum/Simplexcel/issues/12): Sanitizing Regex stripped out too many characters (like the Ampersand or Emojis). Note that certain Unicode characters only work on newer versions of Excel (e.g., Emojis work in Excel 2013 but not 2007 or 2010).

## 2.0.0 (2017-04-22)
* Re-target to .net Framework 4.5 and .NET Standard 1.3.
* No longer use `System.Drawing.Color` but new type `Simplexcel.Color` should work.
* Classes no longer use Data Contract serializer, hence no more `[DataContract]`, `[DataMember]`, etc. attributes.
* Remove `CompressionLevel` - the entire creation of the actual .xlsx file is re-done (no more dependency on `System.IO.Packaging`) and compression is now a simple bool.

## 1.0.5 (2014-01-30)
* SharedStrings are sanitized to avoid XML Errors when using Escape chars (like 0x1B).

## 1.0.4 (2014-01-21)
* Workbook.Save throws an InvalidOperationException if there are no sheets.

## 1.0.3 (2013-08-20)
* Added support for external hyperlinks.
* Made Workbooks serializable using the .net DataContractSerializer.

## 1.0.2 (2013-01-10)
* Initial Public Release.

# License
http://mstum.mit-license.org/

The MIT License (MIT)
 
Copyright (c) 2013-2017 Michael Stum, http://www.Stum.de <opensource@stum.de>  
Contains contributions by [@mguinness](https://github.com/mguinness)

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.