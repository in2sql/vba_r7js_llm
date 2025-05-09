# ListObject Resize Method

## Business Description
The Resize method allows a ListObject object to be resized over a new range. No cells are inserted or moved.

## Behavior
TheResizemethod allows aListObjectobject  to be resized over a new range.  No cells are inserted or moved.

## Example Usage
```vba
Sub ResizeList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 objListObj.ResizeRange("A1:B10") 
End Sub
```