# PivotField DataRange Property

## Business Description
Returns a Range object as shown in the following table. Read-only.

## Behavior
Returns aRangeobject as shown in the following table. Read-only.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Worksheets("Sheet1").Activate 
pvtTable.PivotFields("REGION").DataRange.Select
```