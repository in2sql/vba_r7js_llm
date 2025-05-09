Sub DisplayUserName() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Display name of user with access to protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users(1).Name 
 
End Sub