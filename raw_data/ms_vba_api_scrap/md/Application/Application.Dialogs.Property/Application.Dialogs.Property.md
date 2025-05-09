# Application.Dialogs property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
Application.Dialogs(xlDialogOpen).Show
```

## Example
```vba
Sub SendIt() 
    Application.Dialogs(xlDialogSendMail).Show arg1:="ask@mrexcel.com", arg2:="This goes in the subject line" 
End Sub
```

