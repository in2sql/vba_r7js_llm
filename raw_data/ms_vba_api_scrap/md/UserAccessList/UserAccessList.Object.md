# UserAccessList Object

## Business Description
A collection of UserAccess objects that represent the user access for protected ranges.

## Behavior
A collection ofUserAccessobjects that represent the user access for protected ranges.

## Example Usage
```vba
Sub UseDeleteAll() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Notify the user the number of users that can access the protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users.Count 
 
End Sub
```