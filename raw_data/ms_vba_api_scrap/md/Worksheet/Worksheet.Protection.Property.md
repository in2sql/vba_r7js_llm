# Worksheet Protection Property

## Business Description
Returns a Protection object that represents the protection options of the worksheet.

## Behavior
Returns aProtectionobject that represents the protection options of the worksheet.

## Example Usage
```vba
Sub CheckProtection() 
 
 ActiveSheet.Protect 
 
 ' Check the ability to insert columns on a protected sheet. 
 ' Notify the user of this status. 
 If ActiveSheet.Protection.AllowInsertingColumns = True Then 
 MsgBox "The insertion of columns is allowed on this protected worksheet." 
 Else 
 MsgBox "The insertion of columns is not allowed on this protected worksheet." 
 End If 
 
End Sub
```