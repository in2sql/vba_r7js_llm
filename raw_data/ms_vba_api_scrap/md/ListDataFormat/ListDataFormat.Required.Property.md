# ListDataFormat Required Property

## Business Description
Returns a Boolean value indicating whether the schema definition of a column requires data before the row is committed. Read-only Boolean.

## Behavior
Returns aBooleanvalue indicating whether the schema definition of a column requires data before the row is committed. Read-onlyBoolean.

## Example Usage
```vba
Sub Test() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.RequiredEnd Sub
```