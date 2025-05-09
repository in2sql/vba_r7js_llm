# Worksheet PivotTableBeforeCommitChanges Event

## Business Description
Occurs before changes are committed against the OLAP data source for a PivotTable.

## Behavior
Occurs before changes are committed against the OLAP data source for a PivotTable.

## Example Usage
```vba
Sub Worksheet_PivotTableBeforeCommitChanges(ByVal TargetPivotTable As PivotTable, _ 
 ByVal ValueChangeStart As Long, ByVal ValueChangeEnd As Long, Cancel As Boolean) 
 
 Dim UserChoice As VbMsgBoxResult 
 
 UserChoice = MsgBox("Allow updates to be saved to: " + TargetPivotTable.Name + "?", vbYesNo) 
 If UserChoice = vbNo Then Cancel = True 
End Sub
```