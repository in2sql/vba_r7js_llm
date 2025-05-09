# Chart SetElement Method

## Business Description
Sets chart elements on a chart. Read/write MsoChartElementType.

## Behavior
Sets chart elements on a chart. Read/writeMsoChartElementType.

## Example Usage
```vba
ActiveChart.Axes(xlValue).MajorGridlines.Select 
 ActiveChart.SetElement (msoElementChartTitleCenteredOverlay) 
 ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinor) 
 ActiveChart.Walls.Select 
 Application.CommandBars("Clip Art").Visible = False 
 ActiveChart.SetElement (msoElementChartFloorShow)
```