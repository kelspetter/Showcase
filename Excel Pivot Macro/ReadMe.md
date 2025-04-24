# Excel Pivot Table Creation Macro

This workbook allows the user to describe how a pivot table should look through a series of attributes, specified directly on a worksheet. The attributes are specified as a series of key-value pairs, and some formatting attributes (number formats) can be specified using the formatting of the cell. For repeatability the user can define templates to control the look of the pivot tables.

The reason I created this is to greatly simplify creating pivot tables in Excel, espeically when the pivot tables can change the number of rows / columns. There is functionality available that allows you to place a pivot table underneath another one, and to allow you to add a title (i.e. a cell with formatted text that sits above a pivot table). This means you won't have to worry about a pivot table overlapping another when data is refreshed and it has additional rows.

Refer to the workbook [here](<Excel Pivot Macro.xlsm>) for working examples. The full help for the available attributes is also in this workbook.

The VBA code is available in text form [here](PivotCreator.bas). It can be added to an existing macro-enabled workbook, however you will need to ensure that the required template sheets and references are added.

## Examples
Here are some examples of pivot table templates, definitions, and their resulting pivot tables.

### Pivot Table Templates
![Example Master Template and Layout Template](<images/Pivot Templates.png>)

In this case, there is one master template called `Master Source 1`, which declares the source data and calculated fields. The second template, `Layout 1`, controls the layout, declaring which fields need to be added as filters, rows, columns, and the value fields and how they are aggregated.

### Pivot Table Defitions
![Example Pivot Table Definitions](<images/Pivot Definitions.png>)

The actual pivot definitions are here. The first one uses two `Template` calls to declare the source and layout, sets the `Destination`, and sets which item to display in the pivot's filter via `Show Single Label Field 1`. It also declares its `type` to be a master pivot table, meaning other pivot tables can reference this one to re-use the same pivot cache. A `title` is included.

The second and subsequent definitions use the layout template, but specify the `Source Cache` attribute to reference the first pivot table. The filter is set via `Show Single Label Field 1`, as is the `title`. `Place Under` references the previously created pivot table to ensure it will be consistently placed below without overlapping.

### Pivot Table Definition Output
![Example Pivot Table Definition Output](<images/Pivot Definition Output.png>)

The pivot tables in this case are created on a separate sheet, Output.

## Help
The help is part of the [Excel Workbook](<Excel Pivot Macro.xlsm>). Below is an excerp.
![Help Excerp](<Help Excerp.png>)
