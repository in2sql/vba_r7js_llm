# Workbook SheetPivotTableChangeSync Event

## Business Description
Occurs after changes to a PivotTable.

## Behavior
Occurs after changes to a PivotTable.

## Example Usage
```vba
Private Sub Workbook_SheetPivotTableChangeSync(ByVal Sh As Target, Target As PivotTable) 
 
With Target 
 MsgBox "You performed an operation in the following PivotTable: " & .Name & " on " & Sh.Name 
End With 
 
End Sub
```