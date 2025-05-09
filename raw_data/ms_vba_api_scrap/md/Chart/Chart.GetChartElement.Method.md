# Chart GetChartElement Method

## Business Description
Returns information about the chart element at specified X and Y coordinates. This method is unusual in that you specify values for only the first two arguments.

## Behavior
Returns information about the chart element at specified X and Y coordinates. This method is unusual in that you specify values for only the first two arguments. Microsoft Excel fills in the other arguments, and your code should examine those values when the method returns.

## Example Usage
```vba
Private Sub Chart_MouseMove(ByVal Button As Long, _ 
 ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) 
 Dim IDNum As Long 
 Dim a As Long 
 Dim b As Long 
 
 ActiveChart.GetChartElementX, Y, IDNum, a, b 
 If IDNum = xlLegendEntry Then _ 
 MsgBox "WARNING: Move away from the legend" 
End Sub
```