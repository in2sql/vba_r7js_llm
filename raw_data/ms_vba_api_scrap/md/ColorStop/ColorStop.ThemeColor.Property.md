# ColorStop ThemeColor Property

## Business Description
Returns or sets the theme color of the represented object. Read/write

## Behavior
Returns or sets the theme color of the represented object. Read/write

## Example Usage
```vba
Range("A1:A10").Select 
With Selection.Interior.Gradient.ColorStop.Add(1) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With
```