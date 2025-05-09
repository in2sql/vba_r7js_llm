# PlotArea Object

## Business Description
Represents the plot area of a chart.

## Behavior
Represents the plot area of a chart.

## Example Usage
```vba
Charts("Chart1").Activate 
With ActiveChart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
End With
```