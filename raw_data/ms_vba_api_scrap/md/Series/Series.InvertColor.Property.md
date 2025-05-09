# Series InvertColor Property

## Business Description
Returns or sets the fill color for negative data points in a series. Read/write

## Behavior
Returns or sets the fill color for negative data points in a series. Read/write

## Example Usage
```vba
ActiveSheet.ChartObjects("Chart 2").Activate 
ActiveChart.SeriesCollection(1).InvertIfNegative = True 
ActiveChart.SeriesCollection(1).InvertColor = RGB(255, 0, 255)
```