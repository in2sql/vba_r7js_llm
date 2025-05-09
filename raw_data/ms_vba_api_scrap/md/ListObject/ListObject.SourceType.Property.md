# ListObject SourceType Property

## Business Description
Returns a XlListObjectSourceType value that represents the current source of the list.

## Behavior
Returns aXlListObjectSourceTypevalue that represents the current source of the list.

## Example Usage
```vba
Sub Test () 
 Dim wrksht As Worksheet 
 Dim oListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects(1) 
 
 Debug.Print oListObj.SourceTypeEnd Sub
```