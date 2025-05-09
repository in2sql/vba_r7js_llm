# DataLabels ShowValue Property

## Business Description
Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide. Read/write.

## Behavior
Returns or sets aBooleancorresponding to a specified chart's data label values display behavior.Truedisplays the values.Falseto hide. Read/write.

## Example Usage
```vba
Sub UseValue() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowValue= True 
 
End Sub
```