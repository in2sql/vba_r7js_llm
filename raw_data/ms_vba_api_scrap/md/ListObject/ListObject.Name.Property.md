# ListObject Name Property

## Business Description
Returns or sets a String value that represents the name of the ListObject object.

## Behavior
Returns or sets aStringvalue that represents the name of theListObjectobject.

## Example Usage
```vba
Sub Test 
 Dim wrksht As Worksheet 
 Dim oListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects(1) 
 
 MsgBox oListObj.Name 
End Sub
```