# ListObject HeaderRowRange Property

## Business Description
Returns a Range object that represents the range of the header row for a list. Read-only Range.

## Behavior
Returns aRangeobject that represents the range of the header row for a list. Read-onlyRange.

## Example Usage
```vba
Sub ActivateHeaderRow() 
 Dim wrksht As Worksheet 
 Dim objList As ListObject 
 Dim objListRng As Range 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objList = wrksht.ListObjects(1) 
 Set objListRng = objList.HeaderRowRangeobjListRng.Activate 
End Sub
```