# AutoFilter.Filters property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
With Worksheets("Crew") 
 If .AutoFilterMode Then 
 With .AutoFilter.Filters(1) 
 If .On Then c1 = .Criteria1 
 End With 
 End If 
End With
```

## Example
No VBA example available.
