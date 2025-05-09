# Application.ErrorCheckingOptions property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckTextDate() 
 
 ' Enable Microsoft Excel to identify dates written as text. 
 Application.ErrorCheckingOptions.TextDate = True 
 
 Range("A1").Formula = "'April 23, 00" 
 
End Sub
```

## Example
```vba
Sub CheckTextDate() 
 
 ' Enable Microsoft Excel to identify dates written as text. 
 Application.ErrorCheckingOptions.TextDate = True 
 
 Range("A1").Formula = "'April 23, 00" 
 
End Sub
```

