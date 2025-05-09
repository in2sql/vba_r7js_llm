# PivotTable RefreshName Property

## Business Description
Returns the name of the person who last refreshed the PivotTable report data. Read-only String.

## Behavior
Returns the name of the person who last refreshed the PivotTable report data. Read-onlyString.

## Example Usage
```vba
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
MsgBox "The data was last refreshed by " & pvtTable.RefreshName
```