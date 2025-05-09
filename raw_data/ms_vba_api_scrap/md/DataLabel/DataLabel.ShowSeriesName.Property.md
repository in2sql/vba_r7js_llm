# DataLabel ShowSeriesName Property

## Business Description
Returns or sets a Boolean to indicate the series name display behavior for the data labels on a chart. True to show the series name. False to hide. Read/write.

## Behavior
Returns or sets aBooleanto indicate the series name display behavior for the data labels on a chart.Trueto show the series name.Falseto hide. Read/write.

## Example Usage
```vba
Sub UseSeriesName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowSeriesName= True 
 
End Sub
```