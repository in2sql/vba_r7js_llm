# Chart MouseUp Event

## Business Description
Occurs when a mouse button is released while the pointer is over a chart.

## Behavior
Occurs when a mouse button is released while the pointer is over a chart.

## Example Usage
```vba
Private Sub Chart_MouseUp(ByVal Button As Long, _ 
 ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) 
 MsgBox "Button = " & Button & chr$(13) & _ 
 "Shift = " & Shift & chr$(13) & _ 
 "X = " & X & " Y = " & Y 
End Sub
```