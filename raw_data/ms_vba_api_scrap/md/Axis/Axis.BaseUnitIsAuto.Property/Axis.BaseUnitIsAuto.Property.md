# Axis.BaseUnitIsAuto property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
With Worksheets(1).ChartObjects(1).Chart 
 With .Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnitIsAuto = True 
 End With 
End With
```

## Remarks
You cannot set this property for a value axis.

## Example
No VBA example available.
