# Application.ConvertFormula method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
inputFormula = "=SUM(R10C2:R15C2)" 
MsgBox Application.ConvertFormula( _ 
 formula:=inputFormula, _ 
 fromReferenceStyle:=xlR1C1, _ 
 toReferenceStyle:=xlA1)
```

## Parameters
- **Formula**: Required
- **FromReferenceStyle**: Required
- **ToReferenceStyle**: Optional
- **ToAbsolute**: Optional
- **RelativeTo**: Optional

## Return Value
Variant

## Remarks
There is a 255 character limit for the formula.

## Example
No VBA example available.
