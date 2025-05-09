# AutoCorrect.DisplayAutoCorrectOptions property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub CheckDisplaySetting() 
 
 'Determine setting and notify user. 
 If Application.AutoCorrect.DisplayAutoCorrectOptions = True Then 
 MsgBox "The AutoCorrect Options button can be displayed." 
 Else 
 MsgBox "The AutoCorrect Options button cannot be displayed." 
 End If 
 
End Sub
```

## Remarks
The DisplayAutoCorrectOptions property is a Microsoft Office-wide setting. Changing this property in Microsoft Excel will affect the other Office applications also.

## Example
```vba
Sub CheckDisplaySetting() 
 
 'Determine setting and notify user. 
 If Application.AutoCorrect.DisplayAutoCorrectOptions = True Then 
 MsgBox "The AutoCorrect Options button can be displayed." 
 Else 
 MsgBox "The AutoCorrect Options button cannot be displayed." 
 End If 
 
End Sub
```

