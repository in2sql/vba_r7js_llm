# Application.CalculateFullRebuild method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseCalculateFullRebuild() 
 
 If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 Application.CalculateFullRebuild 
 End If 
 
End Sub
```

## Remarks
Dependencies are the formulas that depend on other cells. For example, the formula "=A1" depends on cell A1. The CalculateFullRebuild method is similar to re-entering all formulas.

## Example
```vba
Sub UseCalculateFullRebuild() 
 
 If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 Application.CalculateFullRebuild 
 End If 
 
End Sub
```

