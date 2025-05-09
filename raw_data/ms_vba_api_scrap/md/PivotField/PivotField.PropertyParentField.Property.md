# PivotField PropertyParentField Property

## Business Description
Returns a PivotField object representing the field to which the properties in this field pertain.

## Behavior
Returns aPivotFieldobject representing the field to which the properties in this field pertain.

## Example Usage
```vba
Sub CheckParentField() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(4) 
 
 ' Check for member properties and notify user. 
 If pvtField.IsMemberProperty = False Then 
 MsgBox "No member properties present." 
 Else 
 MsgBox "The parent field of the members is: " & _ 
 pvtField.PropertyParentFieldEnd If 
 
End Sub
```