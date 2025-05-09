# ShowCategoryName Property

## Business Description
Allows the user to show the category name for the data labels on a chart. Read/write Boolean.

## Behavior
Allows the user to show the category name for the data labels on a chart. Read/write Boolean.

## Example Usage
```vba
Sub UseCategoryName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowCategoryName= True 
 
End Sub
```