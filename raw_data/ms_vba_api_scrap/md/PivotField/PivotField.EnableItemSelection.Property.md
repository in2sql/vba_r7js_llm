# PivotField EnableItemSelection Property

## Business Description
When set to False, disables the ability to use the field dropdown in the user interface. The default value is True. Read/write Boolean.

## Behavior
When set toFalse, disables the ability to use the field dropdown in the user interface. The default value isTrue. Read/writeBoolean.

## Example Usage
```vba
Sub UseEnableItemSelection() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.RowFields(1) 
 
 ' Determine setting for property and enable if necessary. 
 If pvtField.EnableItemSelection= False Then 
 pvtField.EnableItemSelection= True 
 MsgBox "Item selection enabled for fields." 
 Else 
 MsgBox "Item selection is already enabled for fields." 
 End If 
 
End Sub
```