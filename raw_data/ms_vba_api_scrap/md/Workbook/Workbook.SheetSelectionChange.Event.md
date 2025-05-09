# Workbook SheetSelectionChange Event

## Business Description
Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).

## Behavior
Occurs when the selection changes on any worksheet (doesn't occur if the selection is on a chart sheet).

## Example Usage
```vba
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, _ 
 ByVal Target As Excel.Range) 
 Application.StatusBar = Sh.Name & ":" & Target.Address 
End Sub
```