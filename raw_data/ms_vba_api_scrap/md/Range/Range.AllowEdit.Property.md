# Range AllowEdit Property

## Business Description
Returns a Boolean value that indicates if the range can be edited on a protected worksheet.

## Behavior
Returns aBooleanvalue that indicates if the range can be edited on a protected worksheet.

## Example Usage
```vba
Sub UseAllowEdit() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Protect the worksheet 
 wksOne.Protect 
 
 ' Notify the user about editing cell A1. 
 If wksOne.Range("A1").AllowEdit= True Then 
 MsgBox "Cell A1 can be edited." 
 Else 
 Msgbox "Cell A1 cannot be edited." 
 End If 
 
End Sub
```