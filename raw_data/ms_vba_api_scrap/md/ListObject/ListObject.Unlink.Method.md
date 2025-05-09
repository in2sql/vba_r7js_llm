# ListObject Unlink Method

## Business Description
Removes the link to a Microsoft SharePoint Foundation site from a list. Returns Nothing.

## Behavior
Removes the link to a Microsoft SharePoint Foundation site from a list. ReturnsNothing.

## Example Usage
```vba
Sub UnlinkList() 
 Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 objListObj.UnlinkEnd Sub
```