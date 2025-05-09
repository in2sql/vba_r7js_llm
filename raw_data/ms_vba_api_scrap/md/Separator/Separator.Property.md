# Separator Property

## Business Description
Allows the user to set or return the separator used for the data labels on a chart. Read/write Variant.

## Behavior
Allows the user to set or return the separator used for the data labels on a chart. Read/write Variant.

## Example Usage
```vba
Sub ChangeSeparator() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.Separator= ";" 
 
End Sub
```