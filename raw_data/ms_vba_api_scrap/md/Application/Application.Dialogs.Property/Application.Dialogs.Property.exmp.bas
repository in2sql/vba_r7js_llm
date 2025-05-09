Sub SendIt() 
    Application.Dialogs(xlDialogSendMail).Show arg1:="ask@mrexcel.com", arg2:="This goes in the subject line" 
End Sub