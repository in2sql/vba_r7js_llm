# PivotTable EnableFieldList Property

## Business Description
False to disable the ability to display the field list for the PivotTable. If the field list was already being displayed it disappears. The default value is True. Read/write Boolean.

## Behavior
Falseto disable the ability to display the field list  for the PivotTable. If the field list was already being displayed it disappears. The default value isTrue. Read/writeBoolean.

## Example Usage
```vba
Sub CheckFieldList() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 ' Determine if field list can be displayed. 
 If pvtTable.EnableFieldList= True Then 
 MsgBox "The field list for the PivotTable can be displayed." 
 Else 
 MsgBox "The field list for the PivotTable cannot be displayed." 
 End If 
 
End Sub
```