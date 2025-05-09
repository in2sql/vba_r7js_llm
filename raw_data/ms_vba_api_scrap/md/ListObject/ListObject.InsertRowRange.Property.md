# ListObject InsertRowRange Property

## Business Description
Returns a Range object representing the Insert row, if any, of a specified ListObject object. Read-only Range.

## Behavior
Returns aRangeobject representing the Insert row, if any, of a specifiedListObjectobject. Read-onlyRange.

## Example Usage
```vba
Function ActivateInsertRow() As Boolean 
 
 Dim wrksht As Worksheet 
 Dim objList As ListObject 
 Dim objListRng As Range 
 
 Set wrksht = ActiveWorkbook.Worksheets(1) 
 Set objList = wrksht.ListObjects(1) 
 Set objListRng = objList.InsertRowRangeIf objListRng Is Nothing Then 
 ActivateInsertRow = False 
 Else 
 objListRng.Activate 
 ActivateInsertRow = True 
 End If 
 
End Function
```