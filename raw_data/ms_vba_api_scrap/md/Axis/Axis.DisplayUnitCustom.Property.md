# Axis.DisplayUnitCustom property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
With Charts("Chart1").Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .HasTitle = True 
 .AxisTitle.Caption = "Rebate Amounts" 
End With
```

## Remarks
Using unit labels when charting large values makes your tick mark labels easier to read. For example, if you label your value axis in units of hundreds, thousands, or millions, you can use smaller numeric values at the tick marks on the axis.

## Example
No VBA example available.
