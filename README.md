# Simplexcel

master: [![Actions Status](https://github.com/mstum/Simplexcel/workflows/CI%20Build/badge.svg?branch=master)](https://github.com/mstum/Simplexcel/actions)

release: [![Actions Status](https://github.com/mstum/Simplexcel/workflows/CI%20Build/badge.svg?branch=release)](https://github.com/mstum/Simplexcel/actions)

This is a simple .xlsx generator library for .net 4.5, .NET Standard 1.3 (or higher), and .NET Standard 2.0.

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

# Usage
See [USAGE.md](https://github.com/mstum/Simplexcel/blob/master/USAGE.md) for instructions how to use.

# Changelog
## 2.2.1 (2018-09-19)
* Fixed bug where Background Color wasn't correctly applied to a Fill. ([Issue 23](https://github.com/mstum/Simplexcel/issues/23))

## 2.2.0 (2018-02-24)
* Add `IgnoredErrors` to a `Cell`, to disable Excel warnings (like "Number stored as text").
* If `LargeNumberHandlingMode.StoreAsText` is set on a sheet, the "Number stored as Text" warning is automatically disabled for that cell.
* Add `Cell.Fill` property, which allows setting the Fill of the cell, including the background color, pattern type (diagonal, crosshatch, grid, etc.) and pattern color
* Add `netstandard2.0` version, on top of the existing `netstandard1.3` and `net45` versions.

## 2.1.0 (2017-09-25)
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
