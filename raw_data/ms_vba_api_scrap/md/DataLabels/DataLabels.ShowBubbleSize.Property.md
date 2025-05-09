# DataLabels ShowBubbleSize Property

## Business Description
True to show the bubble size for the data labels on a chart. False to hide. Read/write Boolean.

## Behavior
Trueto show the bubble size for the data labels on a chart.Falseto hide. Read/writeBoolean.

## Example Usage
```vba
Sub UseBubbleSize() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowBubbleSize= True 
 
End Sub
```