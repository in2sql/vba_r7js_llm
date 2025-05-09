# PivotTable RowGrand Property

## Business Description
True if the PivotTable report shows grand totals for rows. Read/write Boolean.

## Behavior
Trueif the PivotTable report shows grand totals for rows. Read/writeBoolean.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.RowGrand= True
```