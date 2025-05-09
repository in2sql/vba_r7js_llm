# Worksheet PivotTableBeforeAllocateChanges Event

## Business Description
Occurs before changes are applied to a PivotTable.

## Behavior
Occurs before changes are applied to a PivotTable.

## Example Usage
```vba
Sub Worksheet_PivotTableBeforeAllocateChanges(ByVal TargetPivotTable As PivotTable, _ 
 ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean) 
 Dim UserChoice As VbMsgBoxResult 
 
 UserChoice = MsgBox("Allow updates to be applied to: " + TargetPivotTable.Name + "?", vbYesNo) 
 If UserChoice = vbNo Then Cancel = True 
End Sub
```