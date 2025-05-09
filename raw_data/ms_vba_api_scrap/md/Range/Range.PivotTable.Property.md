# Range PivotTable Property

## Business Description
Returns a PivotTable object that represents the PivotTable report containing the upper-left corner of the specified range.

## Behavior
Returns aPivotTableobject that represents the PivotTable report containing the upper-left corner of the specified range.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTablepvtTable.PivotFields("Country").CurrentPage = "Canada"
```