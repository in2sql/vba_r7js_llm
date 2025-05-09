# AllowEditRange object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
Use the Add method or the Item property of the AllowEditRanges collection to return an AllowEditRange object.

## Example
```vba
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 Dim wksPassword As String 
 
 Set wksOne = Application.ActiveSheet 
 
 wksPassword = InputBox ("Enter password for the worksheet") 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=wksPassword 
 
 MsgBox "Cells A1 to A4 can be edited on the protected worksheet." 
 
 ' Change the password. 
 
 wksPassword = InputBox ("Enter the new password for the worksheet") 
 
 wksOne.Protection.AllowEditRanges(1).ChangePassword _ 
 Password:=wksPassword 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```

