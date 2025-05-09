# AllowEditRanges object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
Use the AllowEditRanges property of the Protection object to return an AllowEditRanges collection.

## Example
```vba
Sub UseAllowEditRanges() 
 
 Dim wksOne As Worksheet 
 Dim wksPassword As String 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Unprotect worksheet. 
 wksOne.Unprotect 
 
 wksPassword = InputBox ("Enter password for the worksheet") 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=wksPassword 
 
 ' Notify the user 
 ' the title and address of the range. 
 With wksOne.Protection.AllowEditRanges.Item(1) 
 MsgBox "Title of range: " & .Title 
 MsgBox "Address of range: " & .Range.Address 
 End With 
 
End Sub
```

