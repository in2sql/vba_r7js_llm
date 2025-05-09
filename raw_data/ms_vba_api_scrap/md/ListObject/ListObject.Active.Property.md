# ListObject Active Property

## Business Description
Returns a Boolean value indicating whether a ListObject object in a worksheet is active—that is, whether the active cell is inside the range of the ListObject object. Read-only Boolean.

## Behavior
Returns aBooleanvalue indicating whether aListObjectobject in a worksheet is active—that is, whether the active cell is inside the range of theListObjectobject. Read-onlyBoolean.

## Example Usage
```vba
Function MakeListActive() As Boolean 
 Dim wrksht As Worksheet 
 Dim objList As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objList = wrksht.ListObjects(1) 
 objList.Range.Activate 
 
 MakeListActive = objList.ActiveEnd Function
```