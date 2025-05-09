# UserAccess Object

## Business Description
Represents the user access for a protected range.

## Behavior
Represents the user access for a protected range.

## Example Usage
```vba
Sub UseAllowEditRanges() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Add a range that can be edited on the protected worksheet. 
 wksSheet.Protection.AllowEditRanges.Add "Test", Range("A1") 
 
 ' Notify the user the title of the range that can be edited. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Title 
 
End Sub
```