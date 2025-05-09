# DisplayAutoCorrectOptions Property

## Business Description
Allows the user to display or hide the AutoCorrect Options button. The default value is True. Read/write Boolean.

## Behavior
Allows the user to display or hide the AutoCorrect Options button. The default value is True. Read/write Boolean.

## Example Usage
```vba
Sub CheckDisplaySetting() 
 
 'Determine setting and notify user. 
 If Application.AutoCorrect.DisplayAutoCorrectOptions= True Then 
 MsgBox "The AutoCorrect Options button can be displayed." 
 Else 
 MsgBox "The AutoCorrect Options button cannot be displayed." 
 End If 
 
End Sub
```