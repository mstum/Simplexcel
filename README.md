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

# Not supported
* Automatic Column Sizing
* Merging Cells
* Explicit support for Dates (you can use a custom Cell Format string though)
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
Cells can be created implicitly from strings, int64's, double and decimal (and anything that's implicitly convertible to those, e.g. int32), or explicitly through the Cell constructor.

```cs
Cell textCell = "fromString";
Cell intCell = 4; // will be formatted as number without decimal places "0"
Cell doubleCell = 123.456; // will be formatted with 2 decimal places, "0.00"
Cell decimalCell = 987.654321m; // will be formatted with 2 decimal places, "0.00"
decimalCell.Format = "0.000"; // override the cell format with an Excel Formatting string

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
myCell.TextColor = System.Drawing.Color.Violet;
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
sheet.Cells["A2"].TextColor = System.Drawing.Color.Magenta;
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
        workbook.Save(context.HttpContext.Response.OutputStream, CompressionLevel.NoCompression);
    }
}
```

# Changelog
## 2.0.0 (2017-04-22)
* Re-target to .net Framework 4.5 and .NET Standard 1.3
* No longer use `System.Drawing.Color` but new type `Simplexcel.Color` should work
* Classes no longer use Data Contract serializer, hence no more `[DataContract]`, `[DataMember]`, etc. attributes
* Remove `CompressionLevel` - the entire creation of the actual .xlsx file is re-done (no more dependency on `System.IO.Packaging`) and compression is now a simple bool

## 1.0.5 (2014-01-30)
* SharedStrings are sanitized to avoid XML Errors when using Escape chars (like 0x1B)

## 1.0.4 (2014-01-21)
* Workbook.Save throws an InvalidOperationException if there are no sheets

## 1.0.3 (2013-08-20)
* Added support for external hyperlinks
* Made Workbooks serializable using the .net DataContractSerializer

## 1.0.2 (2013-01-10)
* Initial Public Release

# License
http://mstum.mit-license.org/

The MIT License (MIT)
 
Copyright (c) 2013-2017 Michael Stum, http://www.Stum.de <opensource@stum.de>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.