Function TestMacro() 
    MsgBox ActiveWorkbook.Name 
End Function 
 
Sub AddUDFToCustomCategory() 
    Application.MacroOptions Macro:="TestMacro", Category:="My Custom Category" 
End Sub