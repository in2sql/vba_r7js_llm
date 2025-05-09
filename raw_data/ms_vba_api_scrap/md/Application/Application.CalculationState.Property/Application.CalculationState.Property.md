# Application.CalculationState property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub StillCalculating() 
 
 If Application.CalculationState = xlDone Then 
 MsgBox "Done" 
 Else 
 MsgBox "Not Done" 
 End If 
 
End Sub
```

## Example
```vba
Sub StillCalculating() 
 
 If Application.CalculationState = xlDone Then 
 MsgBox "Done" 
 Else 
 MsgBox "Not Done" 
 End If 
 
End Sub
```

