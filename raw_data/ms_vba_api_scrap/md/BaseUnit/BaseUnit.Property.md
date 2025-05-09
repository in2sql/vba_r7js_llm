# BaseUnit property (Excel Graph)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
With myChart 
 With .Axes(xlCategory) 
 .CategoryType = xlTimeScale 
 .BaseUnit = xlMonths 
 End With 
End With
```

## Remarks
Setting this property has no visible effect if the CategoryType property for the specified axis is set to xlCategoryScale. The set value is retained, however, and takes effect when the CategoryType property is set to xlTimeScale.

## Example
No VBA example available.
