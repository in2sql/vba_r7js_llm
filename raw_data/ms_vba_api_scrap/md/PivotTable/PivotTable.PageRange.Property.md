# PivotTable PageRange Property

## Business Description
Returns a Range object that represents the range that contains the page area in the PivotTable report. Read-only.

## Behavior
Returns aRangeobject that represents the range that contains the page area in the PivotTable report. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A3").Select 
ActiveCell.PivotTable.PageRange.Select
```