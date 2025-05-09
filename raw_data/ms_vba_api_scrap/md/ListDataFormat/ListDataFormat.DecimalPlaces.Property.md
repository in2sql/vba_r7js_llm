# ListDataFormat DecimalPlaces Property

## Business Description
Returns a Long value that represents the number of decimal places to show for the numbers in the ListColumn object. Read-only Long.

## Behavior
Returns aLongvalue that represents the number of decimal places to show for the numbers in theListColumnobject. Read-onlyLong.

## Example Usage
```vba
Function GetDecimalPlaces() As Long 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 GetDecimalPlaces = objListCol.ListDataFormat.DecimalPlacesEnd Function
```