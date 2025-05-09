# Protection AllowEditRanges Property

## Business Description
Returns an AllowEditRanges object.

## Behavior
Returns anAllowEditRangesobject.

## Example Usage
```vba
Sub UseAllowEditRanges() 
 
 Dim wksOne As Worksheet 
 Dim strPwd1 As String 
 
 Set wksOne = Application.ActiveSheet 
 
 strPwd1 = InputBox("Enter Password") 
 
 ' Unprotect worksheet. 
 wksOne.Unprotect 
 
 ' Establish a range that can allow edits 
 ' on the protected worksheet. 
 wksOne.Protection.AllowEditRanges.Add _ 
 Title:="Classified", _ 
 Range:=Range("A1:A4"), _ 
 Password:=strPwd1 
 
 ' Notify the user 
 ' the title and address of the range. 
 With wksOne.Protection.AllowEditRanges.Item(1) 
 MsgBox "Title of range: " & .Title 
 MsgBox "Address of range: " & .Range.Address 
 End With 
 
End Sub
```