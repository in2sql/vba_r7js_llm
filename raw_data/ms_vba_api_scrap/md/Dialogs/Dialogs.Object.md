# Dialogs Object

## Business Description
A collection of all the Dialog objects in Microsoft Excel.

## Behavior
A collection of all theDialogobjects in Microsoft Excel.

## Example Usage
```vba
Sub SendIt() 
    Application.Dialogs(xlDialogSendMail).Show arg1:="ask@mrexcel.com", arg2:="This goes in the subject line" 
End Sub
```