# Chart Resize Event

## Business Description
Occurs when the chart is resized.

## Behavior
Occurs when the chart is resized.

## Example Usage
```vba
Private Sub myChartClass_Resize() 
 With ActiveChart.Parent 
 .Left = 100 
 .Top = 150 
 End With 
End Sub
```