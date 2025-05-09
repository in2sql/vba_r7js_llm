# Error Ignore Property

## Business Description
Allows the user to set or return the state of an error checking option for a range. False enables an error checking option for a range. True disables an error checking option for a range. Read/write Boolean.

## Behavior
Allows the user to set or return the state of an error checking option for a range.Falseenables an error checking option for a range.Truedisables an error checking option for a range. Read/writeBoolean.

## Example Usage
```vba
Sub IgnoreChecking() 
 
 Range("A1").Select 
 
 ' Determine if empty cell references error checking is on, if not turn it on. 
 If Application.Range("A1").Errors(xlEmptyCellReferences).Ignore= True Then 
 Application.Range("A1").Errors(xlEmptyCellReferences).Ignore= False 
 MsgBox "Empty cell references error checking has been enabled for cell A1." 
 Else 
 MsgBox "Empty cell references error checking is already enabled for cell A1." 
 End If 
 
End Sub
```