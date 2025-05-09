# Application.CalculationInterruptKey property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckInterruptKey() 
 
 ' Determine the calculation interrupt key and notify the user. 
 Select Case Application.CalculationInterruptKey 
 Case xlAnyKey 
 MsgBox "The calculation interrupt key is set to any key." 
 Case xlEscKey 
 MsgBox "The calculation interrupt key is set to 'Escape'" 
 Case xlNoKey 
 MsgBox "The calculation interrupt key is set to no key." 
 End Select 
 
End Sub
```

## Example
```vba
Sub CheckInterruptKey() 
 
 ' Determine the calculation interrupt key and notify the user. 
 Select Case Application.CalculationInterruptKey 
 Case xlAnyKey 
 MsgBox "The calculation interrupt key is set to any key." 
 Case xlEscKey 
 MsgBox "The calculation interrupt key is set to 'Escape'" 
 Case xlNoKey 
 MsgBox "The calculation interrupt key is set to no key." 
 End Select 
 
End Sub
```

