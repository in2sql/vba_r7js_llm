Sub UseMapPaperSize() 
 
 ' Determine setting and notify user. 
 If Application.MapPaperSize = True Then 
 MsgBox "Microsoft Excel automatically " & _ 
 "adjusts the paper size according to the country/region setting." 
 Else 
 MsgBox "Microsoft Excel does not " & _ 
 "automatically adjusts the paper size according to the country/region setting." 
 End If 
 
End Sub