# Excel Chart Creation Macro

This workbook allows the user to describe how a chart should look through a series of attributes, specified directly on a worksheet. The attributes are specified as a series of key-value pairs, and some formatting attributes (fonts, colours) can be specified using the formatting of the cell.  For repeatability the user can define templates to control the look of the charts, series, and even individual points.

The reason I created this is to greatly simplify creating charts in Excel. Normally, if you want to have a number of charts with the same look and feel, you would create one chart first, which would then be copy-pasted and data source adjusted, etc. However, this is not always a straight-forward process, as quite often when the data source is changed, the formatting is lost, so individual series need to be reformatted each time.

Refer to the workbook [here](<Excel Chart Macro.xlsm>) for working examples. The full help for the available attributes is also in this workbook.

The VBA code is available in text form [here](ChartCreator.bas). It can be added to an existing macro-enabled workbook, however you will need to ensure that the required template sheets and references are added.

## Examples
Here are some examples of chart definitions and their resulting charts. Note that these refer to templates, and the template declarations are below as well.

### Column and Line Chart
![Example Column and Line Chart](<Example Column and Line Chart.png>)

### Column and Line Chart with additional formatting
![Example Column and Line Chart with Additional Formatting](<Example Column Line Chart Additional Formatting.png>)

### Pie Chart
![Example Pie Chart](<Example Pie Chart.png>)

### 3D Pie Chart
![Example 3D Pie Chart](<Example 3D Pie Chart.png>)

### Chart Templates
![Example Chart Template](<Example Chart Templates.png>)

### Series Templates - Columns
![Example Series Template Columns](<Example Series Template Columns.png>)

### Series Templates - Lines
![Example Series Template Lines](<Example Series Template Lines.png>)

### Point Templates
![Example Point Templates](<Example Point Templates.png>)

## Help
The help is part of the [Excel Workbook](<Excel Chart Macro.xlsm>). Below is an excerp.
![Help Excerp](<Help Excerp.png>)
