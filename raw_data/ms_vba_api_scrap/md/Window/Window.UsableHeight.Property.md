# Window UsableHeight Property

## Business Description
Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-only Double.

## Behavior
Returns the maximum height of the space that a window can occupy in the application window area, in points. Read-onlyDouble.

## Example Usage
```vba
With ActiveWindow 
 .WindowState = xlNormal 
 .Top = 1 
 .Left = 1 
 .Height = Application.UsableHeight.Width = Application.UsableWidth 
End With
```