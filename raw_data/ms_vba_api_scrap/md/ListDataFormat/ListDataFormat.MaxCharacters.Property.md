# ListDataFormat MaxCharacters Property

## Business Description
Returns a Long containing the maximum number of characters allowed in the ListColumn object if the Type property is set to xlListDataTypeText or xlListDataTypeMultiLineText. Read-only Long.

## Behavior
Returns aLongcontaining the maximum number of characters allowed in theListColumnobject if theTypeproperty is set toxlListDataTypeTextorxlListDataTypeMultiLineText.  Read-onlyLong.

## Example Usage
```vba
Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.MaxCharacters
```