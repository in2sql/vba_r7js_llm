# ColorStops Clear Method

## Business Description
Clears the represented object.

## Behavior
Clears the represented object.

## Example Usage
```vba
Range("A1:A10").Select 
With Selection.Interior 
 .Pattern = xlPatternLinearGradient 
 .Gradient.Degree = 90 
 .Gradient.ColorStops.Clear 
End With
```