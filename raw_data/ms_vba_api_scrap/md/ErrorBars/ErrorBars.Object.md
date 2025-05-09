# ErrorBars Object

## Business Description
Represents the error bars on a chart series.

## Behavior
Represents the error bars on a chart series.

## Example Usage
```vba
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.SeriesCollection(1).HasErrorBars = True 
ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
```