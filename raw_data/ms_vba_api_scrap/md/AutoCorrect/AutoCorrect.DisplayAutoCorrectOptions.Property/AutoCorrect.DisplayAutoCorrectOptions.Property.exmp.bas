Sub CheckDisplaySetting() 
 
 'Determine setting and notify user. 
 If Application.AutoCorrect.DisplayAutoCorrectOptions = True Then 
 MsgBox "The AutoCorrect Options button can be displayed." 
 Else 
 MsgBox "The AutoCorrect Options button cannot be displayed." 
 End If 
 
End Sub