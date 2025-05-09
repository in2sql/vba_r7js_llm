Sub ChangeSystemSeparators() 
 
 Range("A1").Formula = "1,234,567.89" 
 MsgBox "The system separators will now change." 
 
 ' Define separators and apply. 
 Application.DecimalSeparator = "-" 
 Application.ThousandsSeparator = "-" 
 Application.UseSystemSeparators = False 
 
End Sub