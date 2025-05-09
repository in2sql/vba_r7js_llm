# Window PointsToScreenPixelsX Method

## Business Description
Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a Long value.

## Behavior
Converts a horizontal measurement from points (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as aLongvalue.

## Example Usage
```vba
With ActiveWindow 
 lWinWidth = _.PointsToScreenPixelsX(.Selection.Width) 
 lWinHeight = _ 
 .PointsToScreenPixelsY(.Selection.Height) 
End With
```