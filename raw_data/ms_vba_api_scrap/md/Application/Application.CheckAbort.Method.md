# Application.CheckAbort method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseCheckAbort() 
 
 Dim rngSubtotal As Variant 
 Set rngSubtotal = Application.Range("A10") 
 
 ' Stop recalculation except for designated cell. 
 Application.CheckAbort KeepAbort:=rngSubtotal 
 
End Sub
```

## Parameters
- **KeepAbort**: Optional

## Example
```vba
Sub UseCheckAbort() 
 
 Dim rngSubtotal As Variant 
 Set rngSubtotal = Application.Range("A10") 
 
 ' Stop recalculation except for designated cell. 
 Application.CheckAbort KeepAbort:=rngSubtotal 
 
End Sub
```

