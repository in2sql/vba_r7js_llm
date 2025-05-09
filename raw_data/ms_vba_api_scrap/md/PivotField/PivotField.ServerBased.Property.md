# PivotField ServerBased Property

## Business Description
True if the data source for the specified PivotTable report is external and only the items matching the page field selection are retrieved. Read/write Boolean.

## Behavior
Trueif the data source for the specified PivotTable report is external and only the items matching the page field selection are retrieved. Read/writeBoolean.

## Example Usage
```vba
For Each fld in ActiveSheet.PivotTables(1).PageFields 
 If fld.ServerBased= True Then 
 r = r + 1 
 Worksheets(2).Cells(r, 1).Value = fld.Name 
 End If 
Next
```