Attribute VB_Name = "bring_in_WF"
Sub bring_in_WF()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlManual
Application.DisplayStatusBar = False

Dim family As String, kbtype As String, color As String, comb As String, last_week As String, count As Integer, FirstWeek As Date
Dim current_month As Integer
Dim total As Long, wk1 As Long
Dim coll As Object
Set week_num = CreateObject("System.Collections.ArrayList")

'Delete "Mapped forecast" and "exception" worksheets if it exist
For Each Worksheet In ThisWorkbook.Worksheets
    If Worksheet.name = "Mapped forecast" Or Worksheet.name = "exception" Or Worksheet.name = "TAM" Then
        Worksheet.Delete
    End If
Next Worksheet

'Get exception.csv file
FileToOpen = ThisWorkbook.path & "\Data source\exception.csv"

'Replace exception worksheet with new one
Set closedBook = Workbooks.Open(FileToOpen)
closedBook.Sheets("exception").Copy After:=ThisWorkbook.Sheets("Pivot")
closedBook.Close savechanges:=False

'Get tam.csv file
FileToOpen = ThisWorkbook.path & "\Data source\tam.csv"

'Replace tam worksheet with new one
Set closedBook = Workbooks.Open(FileToOpen)
closedBook.Sheets("tam").Copy After:=ThisWorkbook.Sheets("Pivot")
closedBook.Close savechanges:=False

'Get WF.csv file
FileToOpen = ThisWorkbook.path & "\Data source\WF.csv"

'Replace raw data worksheet with new one
Set closedBook = Workbooks.Open(FileToOpen)
closedBook.Sheets("WF").Copy After:=ThisWorkbook.Sheets("Pivot")
closedBook.Close savechanges:=False

ThisWorkbook.Worksheets("WF").name = "Mapped forecast"
ThisWorkbook.Worksheets("tam").name = "TAM"
Set Forecast = ThisWorkbook.Worksheets("Mapped forecast")
Set exceptional = ThisWorkbook.Worksheets("exception")
Set tam = ThisWorkbook.Worksheets("TAM")
Forecast.Tab.color = RGB(255, 230, 153)
exceptional.Tab.color = RGB(248, 203, 173)
tam.Tab.color = RGB(248, 203, 173)
exceptional_r = Application.WorksheetFunction.CountA(exceptional.Columns(5))
forecast_r = Application.WorksheetFunction.CountA(Forecast.Columns(7))
forecast_c = Application.WorksheetFunction.CountA(Forecast.Rows(1))
exceptional.Columns("A:F").AutoFit

tam_r = Application.WorksheetFunction.CountA(tam.Columns(10))
tam_c = Application.WorksheetFunction.CountA(tam.Rows(1))
tam.ListObjects.Add(xlSrcRange, tam.Range(tam.Cells(1, 1), tam.Cells(tam_r, tam_c)), , xlYes).name = "TAM"

'India/Japan exceptions
For Row = 2 To forecast_r
    family = Forecast.Cells(Row, 6).Value
    kbtype = Forecast.Cells(Row, 18).Value
    color = Forecast.Cells(Row, 19).Value
    comb = family & " " & kbtype & " " & color
    For row2 = 2 To exceptional_r
        If comb = exceptional.Cells(row2, 5).Value Then
            If InStr(Forecast.Cells(Row, 8).Value, "INDIA") > 0 And exceptional.Cells(row2, 6).Value = "INDIA" Then
                Forecast.Cells(Row, 19).Value = color & "_INDIA"
            ElseIf InStr(Forecast.Cells(Row, 8).Value, "JPN2") > 0 And exceptional.Cells(row2, 6).Value = "JP" Then
                Forecast.Cells(Row, 19).Value = color & "_JP"
            End If
        End If
    Next row2
Next Row

''Move APJ Japan forward 6 weeks, IEC TW 4 weeks
For Row = 2 To forecast_r
    If Forecast.Cells(Row, 4).Value = "Inventec Taiwan" Then 'IEC Taiwan
        
        'Find first planning week for that reference date
        For plan_week = 23 To forecast_c
            If Forecast.Cells(Row, 22).Value = Forecast.Cells(1, plan_week).Value Then
                FirstWeek = plan_week
                Exit For
            End If
        Next plan_week
        
        'First week's forecast
        wk1 = Application.WorksheetFunction.Sum(Forecast.Range(Forecast.Cells(Row, FirstWeek), Forecast.Cells(Row, FirstWeek + 2)))
        Forecast.Cells(Row, FirstWeek).Value = wk1

        'pull in by 4 weeks
        Forecast.Range(Forecast.Cells(Row, FirstWeek + 3), Forecast.Cells(Row, forecast_c)).Cut Destination:=Forecast.Range(Forecast.Cells(Row, FirstWeek + 1), Forecast.Cells(Row, forecast_c - 2))

    ElseIf Forecast.Cells(Row, 4).Value = "APJ FUSION - HP JAPAN" Then 'APJ Japan
    
        'Find first planning week for that reference date
        For plan_week = 23 To forecast_c
            If Forecast.Cells(Row, 22).Value = Forecast.Cells(1, plan_week).Value Then
                FirstWeek = plan_week
                Exit For
            End If
        Next plan_week
        
        'First week's forecast
        wk1 = Application.WorksheetFunction.Sum(Forecast.Range(Forecast.Cells(Row, FirstWeek), Forecast.Cells(Row, FirstWeek + 4)))
        Forecast.Cells(Row, FirstWeek).Value = wk1

        'pull in by 6 weeks
        Forecast.Range(Forecast.Cells(Row, FirstWeek + 5), Forecast.Cells(Row, forecast_c)).Cut Destination:=Forecast.Range(Forecast.Cells(Row, FirstWeek + 1), Forecast.Cells(Row, forecast_c - 4))
    End If
Next Row

'Find totals for the next 6 months
If Day(Forecast.Cells(1, 23).Value) > 27 Then
    current_month = Month(Forecast.Cells(1, 23).Value) + 1
Else
    current_month = Month(Forecast.Cells(1, 23).Value)
End If

For col = forecast_c + 1 To forecast_c + 6 'go through 6 months
    If current_month = 13 Then
        current_month = 1
    End If
    Forecast.Cells(1, col).Value = MonthName(current_month, True)

    week_num.Clear
    For week = 23 To forecast_c 'go through all weeks to find how many weeks belong to this month
        If Month(Forecast.Cells(1, week).Value) = current_month And Day(Forecast.Cells(1, week).Value) < 28 Then
            week_num.Add (week)
        ElseIf Month(Forecast.Cells(1, week).Value) = current_month - 1 And Day(Forecast.Cells(1, week).Value) > 27 Then
            week_num.Add (week)
        End If
    Next week
    
    'calculation
    
    If week_num.count = 0 Then
    
        Forecast.Range(Forecast.Cells(2, col), Forecast.Cells(forecast_r, col)).Value = 0
        
    Else
        Forecast.Range(Forecast.Cells(2, col), Forecast.Cells(forecast_r, col)).FormulaR1C1 = "=sum(R[0]C[-" & col - week_num.Item(0) & "]:R[0]C[-" & col - week_num.Item(week_num.count - 1) & "])"
        current_month = current_month + 1
    End If

Next col

'Total of first 3 months
Forecast.Cells(1, forecast_c + 7).Value = "3M"
Forecast.Range(Forecast.Cells(2, forecast_c + 7), Forecast.Cells(forecast_r, forecast_c + 7)).FormulaR1C1 = "=SUM(RC[-6]:RC[-4])"

'Total of all 6 months
Forecast.Cells(1, forecast_c + 8).Value = "Total"
Forecast.Range(Forecast.Cells(2, forecast_c + 8), Forecast.Cells(forecast_r, forecast_c + 8)).FormulaR1C1 = "=SUM(RC[-7]:RC[-2])"

'if there are missing values, turn column header red
Forecast.Range(Forecast.Cells(1, 8), Forecast.Cells(1, 21)).Interior.color = RGB(112, 173, 71)
For col = 8 To 21
    If col <> 19 Then
        For Row = 2 To forecast_r
            If Forecast.Cells(Row, col) = "" Then
                Forecast.Cells(1, col).Interior.color = RGB(192, 0, 0)
                Exit For
            End If
        Next Row
    End If
Next col

'kb color special rules
For Row = 2 To forecast_r
    If Forecast.Cells(Row, 19) = "" Then
        If Forecast.Cells(Row, 5) = "Consumer" Or (Forecast.Cells(Row, 5) = "Commercial" And Forecast.Cells(Row, 14) = "BNB" And Forecast.Cells(Row, 15) = "PROBOOK") Then
            Forecast.Cells(1, 19).Interior.color = RGB(192, 0, 0)
            Exit For
        End If
    End If
Next Row

'turn into table
forecast_c_final = Application.WorksheetFunction.CountA(Forecast.Rows(1))
Forecast.ListObjects.Add(xlSrcRange, Forecast.Range(Forecast.Cells(1, 1), Forecast.Cells(forecast_r, forecast_c_final)), , xlYes).name = "Mapped_Forecast"

'Change Pivot Table Data Source Range Address
Set piv = ThisWorkbook.Worksheets("Pivot")
Set piv_table = piv.PivotTables("Demand_WF")
piv_table.ChangePivotCache _
    ThisWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=Forecast.ListObjects("Mapped_Forecast").Range)

'Refresh pivot table
piv_table.RefreshTable
piv_table.PivotCache.RefreshOnFileOpen = True

'Delete all calculated pivot fields
'For Each calpf In piv_table.CalculatedFields
'    calpf.Delete
'Next calpf

'Remove value fields
For Each pf In piv_table.DataFields
  pf.Orientation = xlHidden
Next pf


'add month totals to pivot table value fields
count = 1
For weeks = (forecast_c + 1) To forecast_c_final '6 months and 3M
    last_week = Forecast.Cells(1, weeks).Value
    piv_table.AddDataField piv_table.PivotFields(last_week), " " & last_week, xlSum
    piv_table.DataPivotField.PivotItems(" " & last_week).Position = count
    piv_table.PivotFields(" " & last_week).NumberFormat = "#,##0_);[Red](#,##0)"
    count = count + 1
Next weeks

'add calculated fields of month total to pivot table value fields
count = 1
For weeks = (forecast_c + 1) To forecast_c_final '6 months and 3M
    last_week = Forecast.Cells(1, weeks).Value
    piv_table.AddDataField piv_table.PivotFields(last_week), last_week & " Diff", xlSum
    
    With piv_table.PivotFields(last_week & " Diff")
        .Calculation = xlDifferenceFrom
        .BaseField = "Reference Date"
        .BaseItem = "(previous)"
        .NumberFormat = "#,##0_);[Red](#,##0)"
    End With

    count = count + 1
Next weeks

'add planning weeks to pivot table value fields
For weeks = 23 To forecast_c
    last_week = Forecast.Cells(1, weeks).Value
    piv_table.AddDataField piv_table.PivotFields(last_week), " " & last_week, xlSum
    piv_table.PivotFields(" " & last_week).NumberFormat = "#,##0_);[Red](#,##0)"
Next weeks


piv.Activate

End Sub




