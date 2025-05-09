# PivotTable ColumnRange Property

## Business Description
Returns a Range object that represents the range that contains the column area in the PivotTable report. Read-only.

## Behavior
Returns aRangeobject that represents the range that contains the column area in the PivotTable report. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.ColumnRange.Select
```