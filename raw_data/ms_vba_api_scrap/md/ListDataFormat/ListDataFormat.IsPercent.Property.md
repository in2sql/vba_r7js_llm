# ListDataFormat IsPercent Property

## Business Description
Returns a Boolean value. Returns True only if the number data for the ListColumn object will be shown in percentage formatting. Read-only Boolean. Read-only.

## Behavior
Returns aBooleanvalue. ReturnsTrueonly if the number data for theListColumnobject will be shown in percentage formatting. Read-onlyBoolean. Read-only.

## Example Usage
```vba
Function GetIsPercent() As Boolean 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 GetIsPercent = objListCol.ListDataFormat.IsPercentEnd Function
```