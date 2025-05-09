# Workbook SheetDeactivate Event

## Business Description
Occurs when any sheet is deactivated.

## Behavior
Occurs when any sheet is deactivated.

## Example Usage
```vba
Private Sub Workbook_SheetDeactivate(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```