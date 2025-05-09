# Application.SpellingOptions property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub MixedDigitCheck() 
 
 ' Determine the setting on spell checking for mixed digits. 
 If Application.SpellingOptions.IgnoreMixedDigits = True Then 
 MsgBox "The spelling options are set to ignore mixed digits." 
 Else 
 MsgBox "The spelling options are set to check for mixed digits." 
 End If 
 
End Sub
```

## Example
```vba
Sub MixedDigitCheck() 
 
 ' Determine the setting on spell checking for mixed digits. 
 If Application.SpellingOptions.IgnoreMixedDigits = True Then 
 MsgBox "The spelling options are set to ignore mixed digits." 
 Else 
 MsgBox "The spelling options are set to check for mixed digits." 
 End If 
 
End Sub
```

