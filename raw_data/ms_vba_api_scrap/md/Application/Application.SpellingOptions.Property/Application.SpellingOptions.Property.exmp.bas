Sub MixedDigitCheck() 
 
 ' Determine the setting on spell checking for mixed digits. 
 If Application.SpellingOptions.IgnoreMixedDigits = True Then 
 MsgBox "The spelling options are set to ignore mixed digits." 
 Else 
 MsgBox "The spelling options are set to check for mixed digits." 
 End If 
 
End Sub