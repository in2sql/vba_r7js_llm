# ListObject ShowTotals Property

## Business Description
Gets or sets a Boolean to indicate whether the Total row is visible. Read/write Boolean.

## Behavior
Gets or sets aBooleanto indicate whether the Total row is visible.  Read/writeBoolean.

## Example Usage
```vba
Sub test() 
Dim oListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects(1) 
 
 Debug.Print oListObj.ShowTotalsEnd Sub
```