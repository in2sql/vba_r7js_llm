# DataBarBorder Color Property

## Business Description
Returns an object that specifies the color of the border of data bars specified by a conditional formatting rule. Read-only

## Behavior
Returns an object that specifies the color of the border of data bars specified by a conditional formatting rule. Read-only

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