Attribute VB_Name = "Dates"
Public Function fexDateToElsDbDateString(ByVal targetDate As Date) As String
    fexDateToElsDbDateString = toyyyyMMdd(targetDate)
End Function

Public Function fexToyyyyMMdd(d As Date) As String
    fexToyyyyMMdd = Replace(CStr(d), "-", "")
End Function


Function fexConvertWeekNumToStartDate(year As Integer, weekNum As Integer, weekday As String)

    Dim adj As Integer
    Dim strInitDate As String:  strInitDate = CStr(year) & "-01-01"
    
    Select Case weekday
        Case "Sun"
            adj = 8
        Case "Mon"
            adj = 2
        Case "Tue"
            adj = 3
        Case "Wed"
            adj = 4
        Case "Thu"
            adj = 5
        Case "Fri"
            adj = 6
        Case "Sat"
            adj = 7
    End Select
       
        
    adj = adj - WorksheetFunction.weekday(DateValue(strInitDate))
    fexConvertWeekNumToStartDate = (weekNum - 1) * 7 + DateValue(strInitDate) + adj
        
End Function
