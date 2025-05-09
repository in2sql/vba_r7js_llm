# FillFormat GradientAngle Property

## Business Description
Returns or sets the angle of the gradient fill for the specified fill format. Read/write

## Behavior
Returns or sets the angle of the gradient fill for the specified fill format. Read/write

## Example Usage
```vba
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.SeriesCollection(1).Format.Fill.GradientAngle = 45
```