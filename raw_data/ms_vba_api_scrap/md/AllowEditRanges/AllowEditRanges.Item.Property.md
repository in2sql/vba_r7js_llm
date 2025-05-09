# AllowEditRanges.Item property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:="secret" 
 
 MsgBox "Cells A1 to A4 can be edited on the protected worksheet." 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges.Item(1).ChangePassword _ 
 Password:="moresecret" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```

## Parameters
- **Index**: Required

## Example
```vba
Sub UseChangePassword() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:="secret" 
 
 MsgBox "Cells A1 to A4 can be edited on the protected worksheet." 
 
 ' Change the password. 
 wksOne.Protection.AllowEditRanges.Item(1).ChangePassword _ 
 Password:="moresecret" 
 
 MsgBox "The password for these cells has been changed." 
 
End Sub
```

