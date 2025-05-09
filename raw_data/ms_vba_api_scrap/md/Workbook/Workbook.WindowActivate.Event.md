# Workbook WindowActivate Event

## Business Description
Occurs when any workbook window is activated.

## Behavior
Occurs when any workbook window is activated.

## Example Usage
```vba
Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window) 
 Wn.WindowState = xlMaximized 
End Sub
```