# PivotTable RowRange Property

## Business Description
Returns a Range object that represents the range including the row area on the PivotTable report. Read-only.

## Behavior
Returns aRangeobject that represents the range including the row area on the PivotTable report. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.RowRange.Select
```