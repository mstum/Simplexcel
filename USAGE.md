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

## Fills (since 2.2.0)
You can set the Fill of a cell, for example:

```cs
// Specify all parameters of a PatternFill
sheet.Cells["C5"].Fill = new PatternFill { PatternType = PatternType.ThinDiagonalCrosshatch, BackgroundColor = Color.Yellow, PatternColor = Color.Violet };

// Setting only a background color will set the PatternColor to be automatically chosen by Excel
sheet.Cells["C7"].Fill = new PatternFill { PatternType = PatternType.DiagonalCrosshatch, BackgroundColor = Color.Violet };

// Only set the background color. This automatically sets the PatternType to Solid the first time.
sheet.Cells["C9"].Fill.BackgroundColor = Color.Navy;
```

The `PatternType` follows the naming convention in Excel:

![Pattern Styles](/doc/patternstyles.png?raw=true "Pattern Styles")

* Solid, Gray750, Gray500, Gray250, Gray125, Gray0625 (That's 75%, 50%, 25%, 12.5% and 6.25% Gray)
* HorizontalStripe, VerticalStripe, ReverseDiagonalStripe, DiagonalStripe, DiagonalCrosshatch, ThickDiagonalCrosshatch
* ThinHorizontalStripe, ThinVerticalStripe, ThinReverseDiagonalStripe, ThinDiagonalStripe, ThinHorizontalCrosshatch, ThinDiagonalCrosshatch