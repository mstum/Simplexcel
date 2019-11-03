# Notes
*(This is my scratchpad, not all features here might actually make it)*

## Chart support
Content Types:
```xml
    <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
    <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
    <Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>
    <Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
```

* sheet contains reference to a drawing: `<drawing r:id="rId1"/>`
* Drawing references the chart, and anchors it on the sheet
* Charts are... complicated. The functionality is vast, but for our purposes, we usually just need "simple" charts with 2 axes.
    * Line Chart
    * Bar Chart
    * Pie Chart
    * Stacked Columns/Bars
    * Stacked Area
* Charts do not require OLE/GDI/Binary stuff, and an Excel 2019 generated chart works fine in Excel 2007

Charts have Relationships to colors and styles. It seems that multiple charts can share styles, but that means modifying one style affects all charts.
Excel uses a separate colors and style file for each chart, even if they are identical.

```xml
    <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>
```