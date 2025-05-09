# PivotField LabelRange Property

## Business Description
Returns a Range object that represents the cell (or cells) that contain the field label. Read-only

## Behavior
Returns aRangeobject that represents the cell (or cells) that contain the field label. Read-only

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Set pvtField = pvtTable.PivotFields("ORDER_DATE") 
Worksheets("Sheet1").Activate 
pvtField.LabelRange.Select
```