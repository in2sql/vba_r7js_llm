# PivotTable EnableDataValueEditing Property

## Business Description
True to disable the alert for when the user overwrites values in the data area of the PivotTable. True also allows the user to change data values that previously could not be changed. The default value is False. Read/write Boolean.

## Behavior
Trueto disable the alert for when the user overwrites values in the data area of the PivotTable.Truealso allows the user to change data values that previously could not be changed. The default value isFalse. Read/writeBoolean.

## Example Usage
```vba
Sub CheckAlertSetting() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine alert setting. 
 If pvtTable.EnableDataValueEditing= False Then 
 MsgBox "Alert is enabled." 
 Else 
 MsgBox "Alert is disabled." 
 End If 
 
End Sub
```