# Workbook SheetActivate Event

## Business Description
Occurs when any sheet is activated.

## Behavior
Occurs when any sheet is activated.

## Example Usage
```vba
Private Sub Workbook_SheetActivate(ByVal Sh As Object) 
 MsgBox Sh.Name 
End Sub
```