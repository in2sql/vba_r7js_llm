# PivotTable RefreshTable Method

## Business Description
Refreshes the PivotTable report from the source data. Returns True if it's successful.

## Behavior
Refreshes the PivotTable report from the source data. ReturnsTrueif it's successful.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.RefreshTable
```