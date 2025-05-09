Sub CheckTextDate() 
 
 ' Enable Microsoft Excel to identify dates written as text. 
 Application.ErrorCheckingOptions.TextDate = True 
 
 Range("A1").Formula = "'April 23, 00" 
 
End Sub