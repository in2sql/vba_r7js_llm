# Chart SeriesChange Event

## Business Description
Occurs when the user changes the value of a chart data point by clicking a bar in the chart and dragging the top edge up or down thus changing the value of the data point.

## Behavior
Occurs when the user changes the value of a chart data point by clicking a bar in the chart and dragging the top edge up or down thus changing the value of the data point.

## Example Usage
```vba
Private Sub Chart_SeriesChange(ByVal SeriesIndex As Long, _ 
 ByVal PointIndex As Long) 
 Set p = Me.SeriesCollection(SeriesIndex).Points(PointIndex) 
 p.Border.ColorIndex = 3 
End Sub
```