# DataLabel ShowCategoryName Property

## Business Description
True to display the category name for the data labels on a chart. False to hide. Read/write Boolean.

## Behavior
Trueto display the category name for the data labels on a chart.Falseto hide. Read/writeBoolean.

## Example Usage
```vba
Sub UseCategoryName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowCategoryName= True 
 
End Sub
```