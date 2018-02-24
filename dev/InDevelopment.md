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
* Charts are... complicated.