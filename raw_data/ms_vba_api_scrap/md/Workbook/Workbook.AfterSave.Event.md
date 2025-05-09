# Workbook AfterSave Event

## Business Description
Occurs after the workbook is saved.

## Behavior
Occurs after the workbook is saved.

## Example Usage
```vba
Private Sub Workbook_AfterSave(ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The workbook was successfully saved.") 
End If 
End Sub
```