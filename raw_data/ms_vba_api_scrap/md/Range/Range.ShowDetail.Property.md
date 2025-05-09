# Range ShowDetail Property

## Business Description
True if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. Read/write Variant.

## Behavior
Trueif the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. Read/writeVariant. For thePivotItemobject (or theRangeobject if the range is in a PivotTable report), this property is set toTrueif the item is showing detail.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Set myRange = ActiveCell.CurrentRegion 
lastRow = myRange.Rows.Count 
myRange.Rows(lastRow).ShowDetail= True
```