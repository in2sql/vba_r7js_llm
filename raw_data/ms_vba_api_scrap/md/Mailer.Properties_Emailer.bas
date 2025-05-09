Attribute VB_Name = "Main"
Option Explicit

'Import the 'Sleep' function
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub MEGA_emailer()
    Application.ScreenUpdating = False
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Email Tester")
    
    Dim outlookApp As Object: Set outlookApp = CreateObject("Outlook.Application")
    Dim outlookMail As Object
    
    'Load user inputs from worksheet
    Dim emailSubject As String: emailSubject = WS.Range("C5").Value2
    Dim filename As String: filename = WS.Range("C9").Value2
    
    'Load emailBody from file located in same folder
    'Use ADODB as it supports UTF-8
    Dim objStream As Object: Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (ThisWorkbook.Path & "/" & filename)
    Dim emailBody As String: emailBody = objStream.ReadText()
    objStream.Close
    
    'Get the row count of column A
    Dim r As Long: r = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Loop through every email address in column A and send them the email
    Dim i As Long
    For i = 2 To r
        '0 for olMailItem
        Set outlookMail = outlookApp.CreateItem(0)
        With outlookMail
            .To = WS.Range("A" & i).Value2
            .Subject = emailSubject
            '1 for olFormatPlain
            '2 for olFormatHTML
            .BodyFormat = 2
            .HTMLBody = emailBody
            .Send
        End With
        'Sleep for 100 milliseconds to avoid looping too quickly
        Sleep (100)
    Next i
    
    Set objStream = Nothing
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    Application.ScreenUpdating = True
    MsgBox "Emails sent!"
End Sub
