# PivotTable RefreshDate Property

## Business Description
Returns the date on which the PivotTable report was last refreshed. Read-only Date.

## Behavior
Returns the date on which the PivotTable report was last refreshed. Read-onlyDate.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
dateString = Format(pvtTable.RefreshDate, "Long Date") 
MsgBox "The data was last refreshed on " & dateString
```