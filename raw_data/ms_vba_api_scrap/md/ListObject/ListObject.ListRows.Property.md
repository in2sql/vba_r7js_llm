# ListObject ListRows Property

## Business Description
Returns a ListRows object that represents all the rows of data in the ListObject object. Read-only.

## Behavior
Returns aListRowsobject that represents all the rows of data in theListObjectobject. Read-only.

## Example Usage
```vba
Sub DeleteListRow(iRowNumber As Integer) 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objListRows As ListRows 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 Set objListRows = objListObj.ListRowsIf (iRowNumber <> 0) And (iRowNumber < objListRows.Count - 1) Then 
 objListRows(iRowNumber).Delete 
 End If 
End Sub
```