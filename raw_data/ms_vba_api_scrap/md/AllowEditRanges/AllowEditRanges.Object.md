# AllowEditRanges Object

## Business Description
A collection of all the AllowEditRange objects that represent the cells that can be edited on a protected worksheet.

## Behavior
A collection of all theAllowEditRangeobjects that represent the cells that can be edited on a protected worksheet.

## Example Usage
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