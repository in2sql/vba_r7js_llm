# ShowSeriesName Property

## Business Description
Allows the user to show the series name for the data labels on a chart. Read/write Boolean.

## Behavior
Allows the user to show the series name for the data labels on a chart. Read/write Boolean.

## Example Usage
```vba
Sub UseSeriesName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowSeriesName= True 
 
End Sub
```