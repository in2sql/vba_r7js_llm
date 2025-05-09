# Workbook WindowResize Event

## Business Description
Occurs when any workbook window is resized.

## Behavior
Occurs when any workbook window is resized.

## Example Usage
```vba
Private Sub Workbook_WindowResize(ByVal Wn As Excel.Window) 
 Application.StatusBar = Wn.Caption & " resized" 
End Sub
```