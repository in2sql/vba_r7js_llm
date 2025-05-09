# Worksheet ChartObjects Method

## Business Description
Returns an object that represents either a single embedded chart (a ChartObject object) or a collection of all the embedded charts (a ChartObjects object) on the sheet.

## Behavior
Returns an object that represents either a single embedded chart (aChartObjectobject) or a collection of all the embedded charts (aChartObjectsobject) on the sheet.

## Example Usage
```vba
Worksheets("Sheet1").ChartObjects(1).Activate 
ActiveChart.SeriesCollection.Add _ 
 source:=Worksheets("Sheet1").Range("B1:B10")
```