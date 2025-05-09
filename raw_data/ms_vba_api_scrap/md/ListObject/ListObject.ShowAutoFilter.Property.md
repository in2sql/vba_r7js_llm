# ListObject ShowAutoFilter Property

## Business Description
Returns Boolean to indicate whether the AutoFilter will be displayed. Read/write Boolean.

## Behavior
ReturnsBooleanto indicate whether the AutoFilter will be displayed. Read/writeBoolean.

## Example Usage
```vba
Dim wrksht As Worksheet 
 Dim oListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListCol = wrksht.ListObjects(1) 
 
 Debug.Print oListCol.ShowAutoFilter
```