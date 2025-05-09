# Application.CalculationVersion property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
If Application.CalculationVersion <> _ 
 Workbooks(1).CalculationVersion Then 
 blnFullCalc = True 
Else 
 blnFullCalc = False 
End If
```

## Remarks
If the workbook was saved in an earlier version of Excel, and if the workbook hasn't been fully recalculated, this property returns 0.

## Example
No VBA example available.
