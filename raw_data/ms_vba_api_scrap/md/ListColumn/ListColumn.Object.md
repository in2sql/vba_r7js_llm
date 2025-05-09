# ListColumn Object

## Business Description
Represents a column in a table.

## Behavior
Represents a column in a table.

## Example Usage
```vba
Sub AddListColumn() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns.Add 
End Sub
```