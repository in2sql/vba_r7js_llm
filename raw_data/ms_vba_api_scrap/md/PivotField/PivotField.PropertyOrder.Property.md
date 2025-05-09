# PivotField PropertyOrder Property

## Business Description
Valid only for PivotTable fields that are member property fields. Returns a Long indicating the display position of the member property within the cube field to which it belongs. Read/write.

## Behavior
Valid only for PivotTable fields that are member property fields. Returns aLongindicating the display position of the member property within the cube field to which it belongs. Read/write.

## Example Usage
```vba
Sub CheckPropertyOrder() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(4) 
 
 ' Check for member properties and notify user. 
 If pvtField.IsMemberProperty = False Then 
 MsgBox "No member properties present." 
 Else 
 MsgBox "The property order of the members is: " & _ 
 pvtField.PropertyOrderEnd If 
 
End Sub
```