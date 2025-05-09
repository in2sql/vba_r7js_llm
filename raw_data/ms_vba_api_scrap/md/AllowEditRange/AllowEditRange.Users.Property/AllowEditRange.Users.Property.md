# AllowEditRange.Users property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub DisplayUserName() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Display name of user with access to protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users(1).Name 
 
End Sub
```

## Example
```vba
Sub DisplayUserName() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Display name of user with access to protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users(1).Name 
 
End Sub
```

