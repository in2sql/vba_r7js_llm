# PivotTable TableRange2 Property

## Business Description
Returns a Range object that represents the range containing the entire PivotTable report, including page fields. Read-only.

## Behavior
Returns aRangeobject that represents the range containing the entire PivotTable report, including page fields. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A3").PivotTable.TableRange2.Select
```