# Chart Object Events

## Business Description
Chart events occur when the user activates or changes a chart. Events on chart sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and select View Code from the shortcut menu.

## Behavior
Chart events occur when the user activates or changes a chart. Events on chart sheets are enabled by default. To view the event procedures for a sheet, right-click the sheet tab and selectView Codefrom the shortcut menu. Select the event name from theProceduredrop-down list box.

## Example Usage
```vba
Private Sub Chart_SeriesChange(ByVal SeriesIndex As Long, _ 
        ByVal PointIndex As Long) 
    Set p = ActiveChart.SeriesCollection(SeriesIndex). _ 
        Points(PointIndex) 
    p.Border.ColorIndex = 3 
End Sub
```