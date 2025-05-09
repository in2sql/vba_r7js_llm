# PivotCell PivotTable Property

## Business Description
Returns a PivotTable object that represents the PivotTable report associated with the PivotCell.

## Behavior
Returns aPivotTableobject that represents the PivotTable report associated with the PivotCell.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTablepvtTable.PivotFields("Country").CurrentPage = "Canada"
```