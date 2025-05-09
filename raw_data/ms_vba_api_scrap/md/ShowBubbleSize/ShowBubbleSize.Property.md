# ShowBubbleSize Property

## Business Description
Allows the user to show the bubble size for the data labels on a chart. Read/write Boolean.

## Behavior
Allows the user to show the bubble size for the data labels on a chart. Read/write Boolean.

## Example Usage
```vba
Sub UseBubbleSize() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowBubbleSize= True 
 
End Sub
```