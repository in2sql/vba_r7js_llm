# Worksheet PivotTableChangeSync Event

## Business Description
Occurs after changes to a PivotTable.

## Behavior
Occurs after changes to a PivotTable.

## Example Usage
```vba
Private Sub Worksheet_PivotTableChangeSync(ByVal Target As PivotTable) 
 
With Target 
 MsgBox "You performed an operation in the following PivotTable: " & .Name 
End With 
 
End Sub
```