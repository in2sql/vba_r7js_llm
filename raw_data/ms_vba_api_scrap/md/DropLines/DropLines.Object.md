# DropLines Object

## Business Description
Represents the drop lines in a chart group.

## Behavior
Represents the drop lines in a chart group.

## Example Usage
```vba
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.ChartGroups(1).HasDropLines = True 
ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
```