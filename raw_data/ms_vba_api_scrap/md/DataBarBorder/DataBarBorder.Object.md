# DataBarBorder Object

## Business Description
Represents the border of the data bars specified by a conditional formatting rule.

## Behavior
Represents the border of the data bars specified by a conditional formatting rule.

## Example Usage
```vba
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
With myDataBar.BarBorder 
 .Type = xlDataBarBorderSolid 
 .Color.ThemeColor = xlThemeColorAccent2 
 .Color.TintAndShade = 0 
End With
```