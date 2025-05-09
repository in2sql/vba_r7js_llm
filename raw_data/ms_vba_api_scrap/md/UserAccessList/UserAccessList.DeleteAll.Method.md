# UserAccessList DeleteAll Method

## Business Description
Removes all users who have access to a protected range on a worksheet.

## Behavior
Removes all users who have access to a protected range on a worksheet.

## Example Usage
```vba
Sub UseDeleteAll() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Remove all users with access to the first protected range. 
 wksSheet.Protection.AllowEditRanges(1).Users.DeleteAll 
 
End Sub
```