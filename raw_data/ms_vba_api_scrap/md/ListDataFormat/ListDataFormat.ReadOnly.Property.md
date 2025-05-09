# ListDataFormat ReadOnly Property

## Business Description
Returns True if the object has been opened as read-only. Read-only Boolean.

## Behavior
ReturnsTrueif the object has been opened as read-only. Read-onlyBoolean.

## Example Usage
```vba
Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.ReadOnly
```