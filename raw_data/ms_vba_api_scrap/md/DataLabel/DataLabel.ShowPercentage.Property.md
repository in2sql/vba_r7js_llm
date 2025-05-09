# DataLabel ShowPercentage Property

## Business Description
True to display the percentage value for the data labels on a chart. False to hide. Read/write Boolean.

## Behavior
Trueto display the percentage value for the data labels on a chart.Falseto hide. Read/writeBoolean.

## Example Usage
```vba
Sub UsePercentage() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowPercentage= True 
 
End Sub
```