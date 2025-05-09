# ListDataFormat lcid Property

## Business Description
Returns a Long value that represents the LCID for the ListColumn object that is specified in the schema definition. Read-only Long.

## Behavior
Returns aLongvalue that represents the LCID for theListColumnobject that is specified in the schema definition. Read-onlyLong.

## Example Usage
```vba
Sub DisplayLCID() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 MsgBox "List LCID: " & objListCol.ListDataFormat.lcid 
End Sub
```