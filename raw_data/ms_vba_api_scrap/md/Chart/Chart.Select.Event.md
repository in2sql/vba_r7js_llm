# Chart Select Event

## Business Description
Occurs when a chart element is selected.

## Behavior
Occurs when a chart element is selected.

## Example Usage
```vba
Private Sub Chart_Select(ByVal ElementID As Long, _ 
 ByVal Arg1 As Long, ByVal Arg2 As Long) 
 If ElementId = xlChartTitle Then 
 MsgBox "please don't change the chart title" 
 End If 
End Sub
```