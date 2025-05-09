# HiLoLines Object

## Business Description
Represents the high-low lines in a chart group.

## Behavior
Represents the high-low lines in a chart group.

## Example Usage
```vba
Worksheets(1).ChartObjects(1).Activate 
ActiveChart.ChartGroups(1).HasHiLoLines = True 
ActiveChart.ChartGroups(1).HiLoLines.Border.Color = RGB(0, 0, 255)
```