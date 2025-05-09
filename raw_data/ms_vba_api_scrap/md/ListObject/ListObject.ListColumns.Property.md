# ListObject ListColumns Property

## Business Description
Returns a ListColumns collection that represents all the columns in a ListObject object. Read-only.

## Behavior
Returns aListColumnscollection that represents all the columns in aListObjectobject. Read-only.

## Example Usage
```vba
Sub DisplayColumnName 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objListCols As ListColumns 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 Set objListCols = oListObj.ListColumnsDebug.Print objListCols(2).Name 
End Sub
```