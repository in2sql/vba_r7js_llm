# ShowValue Property

## Business Description
Allows the user to show the value for the data labels on a chart. Read/write Boolean.

## Behavior
Allows the user to show the value for the data labels on a chart. Read/write Boolean.

## Example Usage
```vba
Sub UseValue() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowValue= True 
 
End Sub
```