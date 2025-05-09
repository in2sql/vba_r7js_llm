# ColorStops Add Method

## Business Description
Adds a ColorStop object to the specified collection.

## Behavior
Adds aColorStopobject to the specified collection.

## Example Usage
```vba
Range("A1:A10").Select 
With Selection.Interior.Gradient.ColorStop.Add(1) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With
```