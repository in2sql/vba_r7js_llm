# Databar BarBorder Property

## Business Description
Returns an object that specifies the border of a data bar. Read-only

## Behavior
Returns an object that specifies the border of a data bar. Read-only

## Example Usage
```vba
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
With myDataBar.BarBorder.Type = xlDataBarBorderSolid 
 .Color.ThemeColor = xlThemeColorAccent2 
 .Color.TintAndShade = 0 
End With
```