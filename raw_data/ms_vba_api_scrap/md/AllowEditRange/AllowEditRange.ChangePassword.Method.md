# AllowEditRange.ChangePassword method (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 Dim strPassword As String 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 
 strPassword = InputBox("Please enter the password for the range") 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=strPassword 
 
 strPassword = InputBox("Please enter the new password for the range") 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges("Classified").ChangePassword _ 
 Password:="strPassword" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```

## Parameters
- **Password**: Required

## Example
```vba
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 Dim strPassword As String 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 
 strPassword = InputBox("Please enter the password for the range") 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=strPassword 
 
 strPassword = InputBox("Please enter the new password for the range") 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges("Classified").ChangePassword _ 
 Password:="strPassword" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```

