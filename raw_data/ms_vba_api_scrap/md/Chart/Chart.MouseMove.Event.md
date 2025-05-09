# Chart MouseMove Event

## Business Description
Occurs when the position of the mouse pointer changes over a chart.

## Behavior
Occurs when the position of the mouse pointer changes over a chart.

## Example Usage
```vba
Private Sub Chart_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) 
 MsgBox "X = " & X & " Y = " & Y 
End Sub
```