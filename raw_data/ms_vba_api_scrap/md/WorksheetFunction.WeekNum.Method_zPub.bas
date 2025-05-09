Attribute VB_Name = "zPub"
Sub YearWeekNumber()
YearWeek = "Year 2021 / CW" & Application.WorksheetFunction.WeekNum(Evaluate("=today()"), 15)
End Sub



