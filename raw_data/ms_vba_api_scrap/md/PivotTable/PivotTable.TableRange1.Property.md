# PivotTable TableRange1 Property

## Business Description
Returns a Range object that represents the range containing the entire PivotTable report, but doesn't include page fields. Read-only.

## Behavior
Returns aRangeobject that represents the range containing the entire PivotTable report, but doesn't include page fields. Read-only.

## Example Usage
```vba
Worksheets("Sheet1").Activate 
Range("A3").PivotTable.TableRange1.Select
```