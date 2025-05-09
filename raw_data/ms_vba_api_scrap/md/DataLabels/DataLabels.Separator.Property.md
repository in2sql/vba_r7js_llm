# DataLabels Separator Property

## Business Description
Sets or returns a Variant representing the separator used for the data labels on a chart. Read/write.

## Behavior
Sets or returns aVariantrepresenting the separator used for the data labels on a chart. Read/write.

## Example Usage
```vba
Sub ChangeSeparator() 
 
 ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1) _ 
 .DataLabels.Separator= ";" 
 
End Sub
```