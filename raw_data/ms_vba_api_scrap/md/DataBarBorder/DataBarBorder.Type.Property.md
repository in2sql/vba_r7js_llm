# DataBarBorder Type Property

## Business Description
Returns or sets the type of border for data bars specified by a conditional formatting rule. Read/write

## Behavior
Returns or sets the type of border for data bars specified by a conditional formatting rule. Read/write

## Example Usage
```vba
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
With myDataBar.BarBorder 
 .Type= xlDataBarBorderSolid 
 .Color.ThemeColor = xlThemeColorAccent2 
 .Color.TintAndShade = 0 
End With
```