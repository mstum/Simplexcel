Simplexcel
==========
This is a simple .xlsx generator library for .net 4.0 Client Profile or higher.  
  
It does not aim to implement the entirety of the Office Open XML Workbook format and all the small and big features Excel offers.  
Instead, it is meant as a way to handle common tasks that can't be handled by other workarounds (e.g., CSV Files or HTML Tables) and is fully supported under ASP.net (unlike, say, COM Interop which Microsoft explicitly doesn't support on a server).

Features
========
* You can store numbers as numbers, so no more unwanted conversion to scientific notation on large numbers!
* You can store text that looks like a number as text, so no more truncation of leading zeroes because Excel thinks it's a number
* You can have multiple Worksheets
* You have basic formatting: Font Name/Size, Bold/Underline/Italic, Color, Border around a cell
* You can specify the size of cells
* Workbooks can be saved compressed or uncompressed (CPU Usage vs. Network Traffic)
* You can specify repeating rows and columns (from the top and left respectively), useful when printing.
* Fully supported in ASP.net and Windows Services
 * Note that I do use System.Drawing.Color. Even though the entire System.Drawing is explicitly unsupported under ASP.net/Windows Services by Microsoft, the Color struct is safe as it has no ties to GDI+.

Not supported
=============
On my ToDo List
---------------
* Automatic Column Sizing
* Merging Cells
* Explicit support for Dates (you can use a custom Cell Format string though)

Unlikely to be ever implemented
-------------------------------
* Embedding any Media or anything OLE
* Creating Charts
* Reading Excel sheets

Untested or Unknown
-------------------
* Mono Support (I use System.IO.Packacing from WindowsBase, which to my understanding works - needs testing!)

Usage
=====
For a full documentation with examples, please go to http://mstum.github.com/Simplexcel/

The Library is available on NuGet and can be installed by searching for Simplexcel or using the NuGet command prompt:

    Install-Package simplexcel

For more information on NuGet see http://www.nuget.org

Changelog
=========
1.0.0
-----
* Initial Release

License
=======
http://mstum.mit-license.org/

The MIT License (MIT)
 
Copyright (c) 2013 Michael Stum, http://www.Stum.de <opensource@stum.de>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
