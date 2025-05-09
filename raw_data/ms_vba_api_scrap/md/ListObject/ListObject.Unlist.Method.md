# ListObject Unlist Method

## Business Description
Removes the list functionality from a ListObject object. After you use this method, the range of cells that made up the the list will be a regular range of data.

## Behavior
Removes the list functionality from aListObjectobject.  After  you use this method, the range of cells that made up the  the list will be a regular range of data.

## Example Usage
```vba
Sub DeList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 objListObj.UnlistEnd Sub
```