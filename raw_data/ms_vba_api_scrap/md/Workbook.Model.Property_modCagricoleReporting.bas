Attribute VB_Name = "modCagricoleReporting"
Option Explicit

'  ********* Columns *********
Private Const TB_COLUMN_EVENT_DATE As String = "Date & time"
Private Const TB_COLUMN_ACTIVITY_CONFIRMATION As String = "Classification"
Private Const TB_COLUMN_RISK_REASON_ID As String = "Reason ID"
Private Const TB_COLUMN_RISK_REASON As String = "Reason"
Private Const TB_COLUMN_RISK_SCORE As String = "Risk score"
Private Const TB_COLUMN_APPLICATION As String = "Application"
Private Const TB_COLUMN_PUID As String = "PUID"
Private Const TB_COLUMN_SESSION As String = "Pinpoint session ID"
Private Const TB_COLUMN_ACTIVITY As String = "Activity"

Private Const COLUMN_WEEK As String = "Week"
Private Const COLUMN_YEAR As String = "Year"
Private Const COLUMN_HELPER_PPSID_WEEK As String = "PPSID_WEEK to Sum"
Private Const COLUMN_HELPER_PUID_WEEK As String = "PUID_WEEK to Sum"

Private Const COLUMN_HELPER_PPSID_WEEK_REASONID As String = "PPSID_WEEK_REASONID to Sum"
Private Const COLUMN_HELPER_PUID_WEEK_REASONID As String = "PUID_WEEK_REASONID to Sum"

Private Const COLUMN_HELPER_PPSID_WEEK_REASONID_CLS As String = "PPSID_WEEK_REASONID_CLS to Sum"
Private Const COLUMN_HELPER_PUID_WEEK_REASONID_CLS As String = "PUID_WEEK_REASONID_CLS to Sum"

Private Const COLUMN_HELPER_PPSID_WEEK_CLS As String = "PPSID_WEEK_CLS to Sum"
Private Const COLUMN_HELPER_PUID_WEEK_CLS As String = "PUID_WEEK_CLS to Sum"

'  ********* End of Columns *********

'  ********* TB Values *********
Private Const TB_CLASSIFICATION_CONFIRMED_FRAUD As String = "confirmed_fraud"
Private Const TB_CLASSIFICATION_CONFIRMED_LEGITIMATE As String = "confirmed_legitimate"
Private Const TB_CLASSIFICATION_UNDETERMINED As String = "undetermined"
Private Const TB_CLASSIFICATION_PENDING As String = "pending_confirmation"
'  ********* End of TB Values *********

Private StrColumnEventDate As String
Private StrColumnActivityConfirmation As String
Private StrColumnRiskReason As String
Private StrColumnRiskReasonId As String
Private StrColumnRiskScore As String
Private StrColumnApplication As String
Private StrColumnPuid As String
Private StrColumnSession As String
Private StrColumnActivity As String

Private strClassificationConfirmedFraud As String
Private strClassificationConfirmedLegitimate As String
Private strClassificationUndetermined As String
Private strClassificationPending As String

Private Const MEASURE_DISTINCT_PUID As String = "Distinct PUID"
Private Const MEASURE_FRAUD_PUID As String = "Fraud PUID"
Private Const MEASURE_LEGITIMATE_PUID As String = "Legitimate PUID"
Private Const MEASURE_PENDING_PUID As String = "Pending PUID"
Private Const MEASURE_UNDETERMINED_PUID As String = "Undetermined PUID"
Private Const MEASURE_TP_RATE_PUID As String = "TP Rate PUID"
Private Const MEASURE_FP_RATE_PUID As String = "FP Rate PUID"

Private Const MEASURE_DISTINCT_SESSION As String = "Distinct Session"
Private Const MEASURE_FRAUD_SESSION As String = "Fraud Session"
Private Const MEASURE_LEGITIMATE_SESSION As String = "Legitimate Session"
Private Const MEASURE_PENDING_SESSION As String = "Pending Session"
Private Const MEASURE_UNDETERMINED_SESSION As String = "Undetermined Session"
Private Const MEASURE_TP_RATE_SESSION As String = "TP Rate SESSION"
Private Const MEASURE_FP_RATE_SESSION As String = "FP Rate SESSION"

Private Const POWERPIVOT_COMADDIN_PROGID As String = "PowerPivotExcelClientAddIn.NativeEntry.1"
Private Const POWERPIVOT_MENUBAR_CONTROL As String = "Power Pivot"

Private Const REASON_ID__2 As String = "-2"
Private Const REASON_ID__1 As String = "-1"
Private Const REASON_ID_19 As String = "19"
Private Const REASON_ID_BLANK As String = "="

Sub cagricole_reporting()
Attribute cagricole_reporting.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Wbk As Workbook
    Dim shtRawData As Worksheet
    Dim strDetectionRateFolderPath As String
    Dim strBoxPath As String
    Dim arrColumnsWithExceptions As Variant
    Dim Pvt As PivotTable
    Const REPORT_NAME As String = "Pivot Table"
    Dim intNumberOfSourceFiles As Integer
    Dim strQueryString As String

    Dim arrVarValuesRiskReasons() As Variant
    
    strBoxPath = Environ("UserProfile") & Application.PathSeparator & "Box" & Application.PathSeparator & "Trusteer\Reporting\cagricole reporting requirements\original\TB export folder containing CSVs"
    strDetectionRateFolderPath = "C:\Users\919561756\Box\Trusteer\Reporting\VBA Projects\FP Monitoring\Cagricole\November 2024"
    If strDetectionRateFolderPath = "False" Then Exit Sub
    Application.ScreenUpdating = False
    intNumberOfSourceFiles = CountFilesInFolder(strDetectionRateFolderPath)
    strQueryString = "let" & Chr(13) & "" & Chr(10) & "    Source = Folder.Files(""" & strDetectionRateFolderPath & """)," & Chr(13) & "" & Chr(10) & "    #""Filtered Hidden Files1"" = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true)," & Chr(13) & "" & Chr(10) & "    #""Invoke Custom Function1"" = Table.AddColumn(#""Filtered Hidden Files1"", ""Transform File"", each #""Trans" & _
        "form File""([Content]))," & Chr(13) & "" & Chr(10) & "    #""Renamed Columns1"" = Table.RenameColumns(#""Invoke Custom Function1"", {""Name"", ""Source.Name""})," & Chr(13) & "" & Chr(10) & "    #""Removed Other Columns1"" = Table.SelectColumns(#""Renamed Columns1"", {""Source.Name"", ""Transform File""})," & Chr(13) & "" & Chr(10) & "    #""Expanded Table Column1"" = Table.ExpandTableColumn(#""Removed Other Columns1"", ""Transform File"", Table.Co" & _
        "lumnNames(#""Transform File""(#""Sample File"")))," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(#""Expanded Table Column1"",{{""Source.Name"", type text}, {""Account Id"", type text}, {""Application"", type text}, {""Browser"", type text}, {""Browser version"", type text}, {""Classification"", type text}, {""Client Language"", type text}, {""Line Carrier"", t" & _
        "ype any}, {""Country code"", type text}, {""Date & time"", type datetime}, {""Customer session IDs"", type text}, {""Device ID"", type text}, {""Encrypted user ID"", type text}, {""City"", type text}, {""Country"", type text}, {""ISP"", type text}, {""IP"", type text}, {""Name"", type text}, {""Machine ID"", type any}, {""Malware Name"", type any}, {""Infected App""" & _
        ", type any}, {""Infected Package"", type any}, {""OS"", type text}, {""Pinpoint session ID"", type text}, {""Platform"", type text}, {""PUID"", type text}, {""Assessment Details"", type text}, {""Recommendation"", type text}, {""Partial result reason"", type any}, {""Reason ID"", Int64.Type}, {""Reason"", type text}, {""Risk score"", Int64.Type}, {""Classified By""," & _
        " type text}, {""Status"", type text}, {""Classified At"", type datetime}, {""New Device"", type logical}, {""Activity"", type text}, {""Closed By"", type any}, {""Closed At"", type any}, {""User Agent"", type text}, {""Assigned To"", type text}, {""Phishing Url"", type any}, {""Detected At"", type text}, {""SDK Configuration"", type any}, {""SDK Version"", type any}" & _
        ", {""MRST App Count"", type any}, {""Call In Progress"", type any}, {""User Behavioral Score"", type any}, {""Risky Device"", type any}, {""Risky Connection"", type any}, {""Battery Charging"", type any}, {""Behavioral Anomaly"", type any}, {""First Seen In Account"", type datetime}, {""First Seen In Region"", type datetime}, {""Fraud MO"", type any}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    
    Set Wbk = Workbooks.Add(xlWBATWorksheet)
    With Wbk
        With .Queries
            If intNumberOfSourceFiles > 1 Then 'if more than 1 source file was found
                .Add name:="foo report name", _
                    formula:=strQueryString
                .Add name:="Sample File", formula:= _
                    "let Source = Folder.Files(""" & strDetectionRateFolderPath & """), Navigation1 = Source{0}[Content] in Navigation1"
                .Add name:="Parameter1", formula:= _
                    "#""Sample File"" meta [IsParameterQuery=true, BinaryIdentifier=#""Sample File"", Type=""Binary"", IsParameterQueryRequired=true]"
                .Add name:="Transform Sample File", formula:= _
                    "let Source = Csv.Document(Parameter1,[Delimiter="","", QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #""Promoted Headers"""
                .Add name:="Transform File", formula:= _
                    "let Source = (Parameter1) => let Source = Csv.Document(Parameter1,[Delimiter="","", QuoteStyle=QuoteStyle.None]), #""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #""Promoted Headers"" in Source"
            Else
                MsgBox "Adjust VBA code to handle importing a single source file"
            End If
        End With
        Set shtRawData = .ActiveSheet
        With shtRawData
            With .ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""foo report name"";Extended Properties=""""" _
            , Destination:=Range("$A$1")).queryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [foo report name]")
            .Refresh BackgroundQuery:=False
            End With
            .name = "Raw Data"
        End With
        
        Call RemoveDuplicates(shtRawData)
        
        arrColumnsWithExceptions = SetDataSourceType(shtRawData)
        If Len(Join(arrColumnsWithExceptions)) = 0 Then
            arrVarValuesRiskReasons = Array(REASON_ID__1, REASON_ID__2, REASON_ID_BLANK)
            Call DeleteIrrelevantRecords(shtRawData, TB_COLUMN_RISK_REASON_ID, arrVarValuesRiskReasons)
        
            Call AddColumns(shtRawData)
            Call CreateNationalReport(shtRawData, REPORT_NAME)
        Else
            MsgBox "The following columns are required for the report:" & vbNewLine & Join(arrColumnsWithExceptions, vbNewLine)
            Exit Sub
        End If
    End With
Application.ScreenUpdating = True
    Set Wbk = Nothing
    Set shtRawData = Nothing
End Sub


'Private Function GetArrayOfMissingColumns(Sht As Worksheet, arrColumns() As Variant) As String()
'    Dim intColumnName As Integer
'    Dim arrColumnsWithExceptions() As String
'    Dim intExceptionCounter As Integer
'
'    For intColumnName = 0 To UBound(arrColumns)
'        If GetSheetColumnIndexByTitle(CStr(arrColumns(intColumnName)), Sht, Sht.Range("A1")) = 0 Then
'            ReDim Preserve arrColumnsWithExceptions(intExceptionCounter)
'            arrColumnsWithExceptions(intExceptionCounter) = CStr(arrColumns(intColumnName))
'            intExceptionCounter = intExceptionCounter + 1
'        End If
'    Next intColumnName
'    GetArrayOfMissingColumns = arrColumnsWithExceptions
'    Erase arrColumns
'End Function



Private Sub RemoveDuplicates(shtRawData As Worksheet)
    Dim intArray As Variant, i As Integer
    Dim rng As Range
    
    Set rng = shtRawData.UsedRange.Rows
    With rng
        ReDim intArray(0 To .Columns.count - 1)
        For i = 0 To UBound(intArray)
            intArray(i) = i + 1
        Next i
        .RemoveDuplicates Columns:=(intArray), Header:=xlYes
    End With
End Sub

Private Sub CreateNationalReport(shtRawData As Worksheet, ReportName As String)
    Dim Pvt As PivotTable
    Dim shtCustomReport As Worksheet
    Dim strColumnSetFormula As String
    Dim pvtField As PivotField
    Dim shtNational As Worksheet
    Dim chartObjAlertEvolution As ChartObject
    Dim rngWeeklyRRSessionCounts As Range
    Dim rngWeeklyRR_TP_RATE_SESSION As Range
    Dim lngRowOffset As Long
    Dim lngColOffset As Long
    Dim Sht As Worksheet
    Dim rngWithCalculatedItem As Range
    
    Set Pvt = GetPivotTable(shtRawData, ReportName) 'Create pivot table
    Set shtCustomReport = ActiveWorkbook.ActiveSheet
    shtCustomReport.name = ReportName

    'Create report "distinct PPSID/PUID by week"
    With Pvt
        .ClearTable
        .ColumnGrand = False
        .RowGrand = False

        With .AddDataField(.PivotFields(COLUMN_HELPER_PPSID_WEEK), "Distinct PPSID", xlSum)
            .NumberFormat = "#,##0"
        End With
        
        With .AddDataField(.PivotFields(COLUMN_HELPER_PUID_WEEK), "Distinct PUID", xlSum)
            .NumberFormat = "#,##0"
        End With
        
        .DataPivotField.Orientation = xlRowField
        
        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Week").Orientation = xlColumnField
        
        Call RemovePivotTableSubtotals(Pvt)
        .RepeatAllLabels xlRepeatLabels
    End With

    Set shtNational = Worksheets.Add
    
    With shtNational
        'Convert values from pivot table to the new worksheet
        Call CopyValues(Pvt.TableRange1, Destination:=.Cells(1))
        
        .Rows(1).Delete
        'Merge cells containing same value
        Call MergeSameCells(Application.Intersect(.UsedRange, .Rows(1)))
    
        Call ChartAlertEvolution(shtNational)
        Set chartObjAlertEvolution = .ChartObjects(.ChartObjects.count)
        
        Set rngWeeklyRRSessionCounts = .Range(chartObjAlertEvolution.BottomRightCell.Address).End(xlToLeft).Offset(1)
        .name = "National"
    End With
    
    With Pvt
        'Reset pivot table
        .ClearTable
        .ColumnGrand = False
        .RowGrand = False

        With .AddDataField(.PivotFields(COLUMN_HELPER_PPSID_WEEK_REASONID), "Distinct PPSID", xlSum)
            .NumberFormat = "#,##0"
        End With
        
        With .AddDataField(.PivotFields(COLUMN_HELPER_PUID_WEEK_REASONID), "Distinct PUID", xlSum)
            .NumberFormat = "#,##0"
        End With
        
        .DataPivotField.Orientation = xlRowField
        
        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Week").Orientation = xlColumnField
        .PivotFields("Reason").Orientation = xlRowField
        
        Call RemovePivotTableSubtotals(Pvt)
        .RepeatAllLabels xlRepeatLabels
    End With
    
    Call CopyValues(Pvt.TableRange1, Destination:=rngWeeklyRRSessionCounts)
    
    With shtNational
        rngWeeklyRRSessionCounts.EntireRow.Delete
        Set rngWeeklyRRSessionCounts = .Cells(.UsedRange.SpecialCells(xlCellTypeLastCell).Row, 1).CurrentRegion
        Call MergeSameCells(rngWeeklyRRSessionCounts.Resize(1, rngWeeklyRRSessionCounts.Columns.count))
        
        With rngWeeklyRRSessionCounts.Resize(rngWeeklyRRSessionCounts.Rows.count, 1).SpecialCells(xlCellTypeBlanks)
            .FormulaR1C1 = "=R[-1]C"
            rngWeeklyRRSessionCounts.Resize(rngWeeklyRRSessionCounts.Rows.count, 1).Value2 = rngWeeklyRRSessionCounts.Resize(rngWeeklyRRSessionCounts.Rows.count, 1).Value2

        End With
        
        Call MergeSameCells(rngWeeklyRRSessionCounts.Resize(rngWeeklyRRSessionCounts.Rows.count, 1))
        
        .UsedRange.EntireColumn.AutoFit
        
        Set rngWeeklyRR_TP_RATE_SESSION = .Cells(.UsedRange.Rows.count, 1).Offset(2)
    End With
    
'''''''''''
    With Pvt
        .ClearTable
        .ColumnGrand = False
        .RowGrand = False

        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Week").Orientation = xlColumnField
        
        With .AddDataField(.PivotFields("Pinpoint session ID"), "Count of Pinpoint session ID", xlCount)
            .NumberFormat = "0%"
        End With
        
        .PivotFields("Classification").Orientation = xlColumnField
        .PivotFields("Reason").Orientation = xlRowField
        .PivotFields(StrColumnApplication).Orientation = xlPageField
        
        .PivotFields("Classification").CalculatedItems.Add "Precision", "=confirmed_fraud / (confirmed_fraud + confirmed_legitimate)", True
        .PivotFields("Classification").CalculatedItems("Precision").StandardFormula = "=confirmed_fraud / (confirmed_fraud + confirmed_legitimate)"
        With .PivotFields("Classification")
            .PivotItems("confirmed_fraud").Visible = False
            .PivotItems("confirmed_legitimate").Visible = False
            .PivotItems("pending_confirmation").Visible = False
            .PivotItems("undetermined").Visible = False
        End With
        Call RemovePivotTableSubtotals(Pvt)
        .RepeatAllLabels xlRepeatLabels
        lngRowOffset = .DataBodyRange.Row - .TableRange1.Row - 1
    End With
    
    Call CopyValues(Pvt.TableRange1.Resize(, Pvt.TableRange1.Columns.count - 1), Destination:=rngWeeklyRR_TP_RATE_SESSION)
    
    With shtNational
        With rngWeeklyRR_TP_RATE_SESSION.Cells(1)
            .Offset(3).EntireRow.Delete
            .EntireRow.Delete
        End With
        
        Set rngWeeklyRR_TP_RATE_SESSION = .Cells(.UsedRange.SpecialCells(xlCellTypeLastCell).Row, 1).CurrentRegion
'Stop
        
        'Implement sub 'RemoveDivisionByZeroColumns' properly
        
        Call MergeSameCells(rngWeeklyRR_TP_RATE_SESSION.Resize(1, rngWeeklyRR_TP_RATE_SESSION.Columns.count))
    End With
    
    Set rngWeeklyRR_TP_RATE_SESSION = rngWeeklyRR_TP_RATE_SESSION.CurrentRegion
    With rngWeeklyRR_TP_RATE_SESSION.Offset(lngRowOffset).Resize(rngWeeklyRR_TP_RATE_SESSION.Rows.count - lngRowOffset)
        .NumberFormat = "0%"
    End With
'''''''''''
    With Pvt
        .ClearTable
        .ColumnGrand = True
        .RowGrand = False

        With .AddDataField(.PivotFields(COLUMN_HELPER_PPSID_WEEK_REASONID), "Distinct PPSID", xlSum)
            .NumberFormat = "#,##0"
        End With
        
        With .AddDataField(.PivotFields(COLUMN_HELPER_PUID_WEEK_REASONID), "Distinct PUID", xlSum)
            .NumberFormat = "#,##0"
        End With
        
        .DataPivotField.Orientation = xlRowField
        
        .PivotFields("Year").Orientation = xlColumnField
        .PivotFields("Week").Orientation = xlColumnField
        .PivotFields("Classification").Orientation = xlRowField
        
        .PivotFields(StrColumnApplication).Orientation = xlPageField

        Call RemovePivotTableSubtotals(Pvt)
        .RepeatAllLabels xlRepeatLabels
        .ShowPages PageField:="Application"
    End With
    
    For Each Sht In ActiveWorkbook.Worksheets
        If Sht.PivotTables.count > 0 Then
            Sht.PivotTables(1).TableRange2.EntireColumn.AutoFit
        End If
    Next Sht
    
    Set Pvt = Nothing
    Set shtCustomReport = Nothing
    Set shtNational = Nothing
End Sub

Private Sub ChartAlertEvolution(shtNational As Worksheet)
    Dim Chrt As Chart
    Dim chartObj As ChartObject
    Dim Fsc As series
    Dim shtPivot As Worksheet
    Dim seriePuid As series
    Dim serieSession As series
    
    With shtNational
        Set Chrt = .Shapes.AddChart2(227, xlLine).Chart
    End With
    
    With Chrt
        For Each Fsc In .FullSeriesCollection
            Fsc.Delete
        Next Fsc
        
        Set seriePuid = .SeriesCollection.NewSeries
        With seriePuid
            .name = "=" & shtNational.name & "!$A$3"
            .Values = "=" & shtNational.name & "!$B$3:$F$3"
        End With
        
        Set serieSession = .SeriesCollection.NewSeries
        With serieSession
            .name = "=" & shtNational.name & "!$A$4"
            .Values = "=" & shtNational.name & "!$B$4:$F$4"
            .XValues = "=" & shtNational.name & "!$B$1:$F$2"
        End With
        
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = "Evolution of alerts"
        
        If .HasLegend Then
            .Legend.Position = xlBottom
        End If
'

        Set chartObj = shtNational.ChartObjects(1)
        With chartObj
            .Left = ActiveWindow.VisibleRange.Range("A1").Left + 2
            .Top = shtNational.Range("A" & shtNational.UsedRange.Rows.count).Offset(1).Top + 2
            .Width = ActiveWindow.VisibleRange.Width * 2 / 3
            .Height = ActiveWindow.VisibleRange.Height / 2
        End With
    End With
        
    Set Chrt = Nothing
    Set Fsc = Nothing
    Set chartObj = Nothing
End Sub

Private Function GetPivotTable(Sht As Worksheet, Optional ReportName As String = "Pivot") As PivotTable
    Dim rngRawData As Range
    Dim shtPivot As Worksheet
    Dim Pvt As PivotTable
    Dim pvtCache As PivotCache
    Dim pvtField As PivotField
    
    Set rngRawData = Sht.Range("A1").CurrentRegion
    With ActiveWorkbook
        Set shtPivot = Worksheets.Add
        
        Set pvtCache = .PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=rngRawData, _
            Version:=6)
        Set Pvt = pvtCache.CreatePivotTable(TableDestination:=shtPivot.Range("A3"), tableName:=ReportName, DefaultVersion:=6)
    End With
    
    'Assign a VB codename to the Pivot Table Worksheet for future References
    Call RenameCodeName(shtPivot, "shtPivot")
    
    With Pvt
        .ColumnGrand = False
        .RowGrand = False
        .InGridDropZones = True
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .NullString = "0"
        .RowAxisLayout xlTabularRow
        For Each pvtField In .PivotFields
            pvtField.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Next pvtField
    End With
    
'    Call AddMeasures
    
    Set GetPivotTable = Pvt

    Set rngRawData = Nothing
    Set shtPivot = Nothing
    Set Pvt = Nothing
    Set pvtCache = Nothing
End Function

Private Sub AddMeasures()
    Dim modelBook As Model
    Dim modelTable As modelTable
    Dim mdlMeasures As ModelMeasures
    Dim formatPercentage As ModelFormatPercentageNumber
    Dim formatWholeNumber As ModelFormatWholeNumber
    Dim formatDecimalNumber As ModelFormatDecimalNumber
    
    Set modelBook = ActiveWorkbook.Model
    Set modelTable = modelBook.ModelTables.item(1)
    Set mdlMeasures = modelBook.ModelMeasures
    
    Set formatPercentage = modelBook.ModelFormatPercentageNumber
    With formatPercentage
        .DecimalPlaces = 0
        .UseThousandSeparator = False
    End With

    Set formatWholeNumber = modelBook.ModelFormatWholeNumber
    With formatWholeNumber
        .UseThousandSeparator = False
    End With
    
    Set formatDecimalNumber = modelBook.ModelFormatDecimalNumber
    With formatDecimalNumber
        .DecimalPlaces = 2
        .UseThousandSeparator = True
    End With
    
    Call AddMeasure(MEASURE_DISTINCT_PUID, "DISTINCTCOUNT(Range[" & StrColumnPuid & "])", formatWholeNumber)
    Call AddMeasure(MEASURE_FRAUD_PUID, "CALCULATE([" & MEASURE_DISTINCT_PUID & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationConfirmedFraud & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_LEGITIMATE_PUID, "CALCULATE([" & MEASURE_DISTINCT_PUID & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationConfirmedLegitimate & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_UNDETERMINED_PUID, "CALCULATE([" & MEASURE_DISTINCT_PUID & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationUndetermined & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_PENDING_PUID, "CALCULATE([" & MEASURE_DISTINCT_PUID & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationPending & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_TP_RATE_PUID, "[" & MEASURE_FRAUD_PUID & "] / ([" & MEASURE_FRAUD_PUID & "] + [" & MEASURE_LEGITIMATE_PUID & "])", formatPercentage)
    Call AddMeasure(MEASURE_FP_RATE_PUID, "[" & MEASURE_LEGITIMATE_PUID & "] / ([" & MEASURE_FRAUD_PUID & "] + [" & MEASURE_LEGITIMATE_PUID & "])", formatPercentage)
    
    Call AddMeasure(MEASURE_DISTINCT_SESSION, "DISTINCTCOUNT(Range[" & StrColumnSession & "])", formatWholeNumber)
    Call AddMeasure(MEASURE_FRAUD_SESSION, "CALCULATE([" & MEASURE_DISTINCT_SESSION & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationConfirmedFraud & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_LEGITIMATE_SESSION, "CALCULATE([" & MEASURE_DISTINCT_SESSION & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationConfirmedLegitimate & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_UNDETERMINED_SESSION, "CALCULATE([" & MEASURE_DISTINCT_SESSION & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationUndetermined & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_PENDING_SESSION, "CALCULATE([" & MEASURE_DISTINCT_SESSION & "],Range[" & StrColumnActivityConfirmation & "]=""" & strClassificationPending & """)", formatWholeNumber)
    Call AddMeasure(MEASURE_TP_RATE_SESSION, "[" & MEASURE_FRAUD_SESSION & "] / ([" & MEASURE_FRAUD_SESSION & "] + [" & MEASURE_LEGITIMATE_SESSION & "])", formatPercentage)
    Call AddMeasure(MEASURE_FP_RATE_SESSION, "[" & MEASURE_LEGITIMATE_SESSION & "] / ([" & MEASURE_FRAUD_SESSION & "] + [" & MEASURE_LEGITIMATE_SESSION & "])", formatPercentage)
    
    Set modelBook = Nothing
    Set modelTable = Nothing
    Set mdlMeasures = Nothing
    Set formatPercentage = Nothing
    Set formatWholeNumber = Nothing
    Set formatDecimalNumber = Nothing
End Sub

Private Function AddMeasure(MeasureName As String, formula As String, FormatInformation As Variant) As ModelMeasure
    Dim intMeasureIndex As Integer
    
    intMeasureIndex = GetMeasureIndex(MeasureName)
    With ActiveWorkbook.Model.ModelMeasures
        If intMeasureIndex = 0 Then
            'the following line may throw an error if powerpivot is not active (despite checked in Com-Addins)
            Set AddMeasure = .Add(MeasureName:=MeasureName, AssociatedTable:=ActiveWorkbook.Model.ModelTables(1), formula:=formula, FormatInformation:=FormatInformation)
        ElseIf intMeasureIndex > 0 Then
            ActiveWorkbook.Model.ModelMeasures.item(MeasureName).formula = formula
        End If
    End With
End Function

Private Function GetMeasureIndex(MeasureName As String) As Integer
    Dim intCount As Integer
    Dim modelBook As Model
    Dim modelTable As modelTable
    
    Set modelBook = ActiveWorkbook.Model
    Set modelTable = modelBook.ModelTables(1)
    For intCount = 1 To modelBook.ModelMeasures.count
        If modelBook.ModelMeasures.item(intCount).name = MeasureName Then
            Exit For
        End If
    Next intCount
    
    If intCount > 0 And intCount <= modelBook.ModelMeasures.count Then
        GetMeasureIndex = intCount
    Else
        GetMeasureIndex = 0
    End If

    Set modelBook = Nothing
    Set modelTable = Nothing
End Function

Private Function getSelectedFolder(Optional OpenAt As Variant) As Variant
    Dim ShellApp As Object
     
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, "Select the folder containing Trustboard .csv export files", 0, OpenAt)
     
    On Error Resume Next
    getSelectedFolder = ShellApp.self.Path
    On Error GoTo 0
     
    Set ShellApp = Nothing
     
     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(getSelectedFolder, 2, 1)
    Case Is = ":"
        If Left(getSelectedFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(getSelectedFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
    
    Set ShellApp = Nothing
    Exit Function
     
Invalid:
     'If it was determined that the selection was invalid, set to False
    getSelectedFolder = False
End Function

Private Sub AddColumns(Sht As Worksheet)
    Call AddColumn(Sht, COLUMN_WEEK, "WEEKNUM(RC[-1], 21)", FormulaPrefix:="W", NumberFormat:="0") 'FormulaR1C1 = "=INT((RC[-1]-SUM(MOD(DATE(YEAR(RC[-1]-MOD(RC[-1]-2,7)+3),1,2),{1E+99;7})*{1;-1})+5)/7)"
    Call AddColumn(Sht, COLUMN_YEAR, "IF(AND(WEEKNUM(RC[-1],21)>50,MONTH(RC[-1])=1),YEAR(RC[-1])-1,YEAR(RC[-1]))", NumberFormat:="0000")

    Call AddColumn(Sht, COLUMN_HELPER_PPSID_WEEK, "1/COUNTIFS([Week],[@Week],[Pinpoint session ID],[@[Pinpoint session ID]])")
    Call AddColumn(Sht, COLUMN_HELPER_PUID_WEEK, "1/COUNTIFS([Week],[@Week],[PUID],[@[PUID]])")

    Call AddColumn(Sht, COLUMN_HELPER_PPSID_WEEK_REASONID, "1/COUNTIFS([Week],[@Week],[Pinpoint session ID],[@[Pinpoint session ID]],[Reason ID],[@[Reason ID]])")
    Call AddColumn(Sht, COLUMN_HELPER_PUID_WEEK_REASONID, "1/COUNTIFS([Week],[@Week],[PUID],[@PUID],[Reason ID],[@[Reason ID]])")

    Call AddColumn(Sht, COLUMN_HELPER_PPSID_WEEK_REASONID_CLS, "1/COUNTIFS([Week],[@Week],[Pinpoint session ID],[@[Pinpoint session ID]],[Reason ID],[@[Reason ID]],[Classification],[@Classification])")
    Call AddColumn(Sht, COLUMN_HELPER_PUID_WEEK_REASONID_CLS, "1/COUNTIFS([Week],[@Week],[PUID],[@PUID],[Reason ID],[@[Reason ID]],[Classification],[@Classification])")

    Call AddColumn(Sht, COLUMN_HELPER_PPSID_WEEK_CLS, "1/COUNTIFS([Week],[@Week],[Pinpoint session ID],[@[Pinpoint session ID]],[Classification],[@Classification])")
    Call AddColumn(Sht, COLUMN_HELPER_PUID_WEEK_CLS, "1/COUNTIFS([Week],[@Week],[PUID],[@PUID],[Classification],[@Classification])")
End Sub

Private Sub AddColumn(Sht As Worksheet, ColumnName As String, FormulaR1C1 As String, Optional FormulaPrefix As String = "", Optional NumberFormat As String)
    Dim rngDataRange As Range
    Dim lngEventDateColIndex As Long
    
    lngEventDateColIndex = GetSheetColumnIndexByTitle(TB_COLUMN_EVENT_DATE, Sht, Sht.Range("A1"))
    Columns(lngEventDateColIndex + 1).Insert
    Cells(1, lngEventDateColIndex + 1).Value2 = ColumnName
    
    Set rngDataRange = GetDataRangeForColumn(Sht, Sht.Range("A1").CurrentRegion, ColumnName)
    With rngDataRange
        If FormulaPrefix <> "" Then
            .FormulaR1C1 = "=CONCAT(""" & FormulaPrefix & """," & FormulaR1C1 & ")"
        Else
            .FormulaR1C1 = "=" & FormulaR1C1
        End If
        .NumberFormat = NumberFormat
        .Value2 = .Value2
    End With

    Set rngDataRange = Nothing
End Sub

'change sub name to getArrColumnsWithExceptions
Private Function SetDataSourceType(Sht As Worksheet) As String()
    Dim arrColumnsWithExceptionsTB() As String
    Dim arrColumnsTB() As Variant
    
    arrColumnsTB = Array(TB_COLUMN_EVENT_DATE, TB_COLUMN_ACTIVITY_CONFIRMATION, TB_COLUMN_RISK_REASON_ID, TB_COLUMN_RISK_REASON, TB_COLUMN_RISK_SCORE, TB_COLUMN_APPLICATION, TB_COLUMN_PUID, TB_COLUMN_SESSION, TB_COLUMN_ACTIVITY)
    arrColumnsWithExceptionsTB = GetArrayOfMissingColumns(Sht, arrColumnsTB)
    If Len(Join(arrColumnsWithExceptionsTB)) = 0 Then
    
        StrColumnEventDate = TB_COLUMN_EVENT_DATE
        StrColumnActivityConfirmation = TB_COLUMN_ACTIVITY_CONFIRMATION
        StrColumnRiskReasonId = TB_COLUMN_RISK_REASON_ID
        StrColumnRiskReason = TB_COLUMN_RISK_REASON
        StrColumnRiskScore = TB_COLUMN_RISK_SCORE
        StrColumnApplication = TB_COLUMN_APPLICATION
        StrColumnPuid = TB_COLUMN_PUID
        StrColumnSession = TB_COLUMN_SESSION
        StrColumnActivity = TB_COLUMN_ACTIVITY
    
        strClassificationConfirmedFraud = TB_CLASSIFICATION_CONFIRMED_FRAUD
        strClassificationConfirmedLegitimate = TB_CLASSIFICATION_CONFIRMED_LEGITIMATE
        strClassificationUndetermined = TB_CLASSIFICATION_UNDETERMINED
        strClassificationPending = TB_CLASSIFICATION_PENDING
    Else
        SetDataSourceType = arrColumnsWithExceptionsTB
    End If
End Function

Private Function GetArrayOfMissingColumns(Sht As Worksheet, arrColumns() As Variant) As String()
    Dim intColumnName As Integer
    Dim arrColumnsWithExceptions() As String
    Dim intExceptionCounter As Integer
    
    For intColumnName = 0 To UBound(arrColumns)
        If GetSheetColumnIndexByTitle(CStr(arrColumns(intColumnName)), Sht, Sht.Range("A1")) = 0 Then
            ReDim Preserve arrColumnsWithExceptions(intExceptionCounter)
            arrColumnsWithExceptions(intExceptionCounter) = CStr(arrColumns(intColumnName))
            intExceptionCounter = intExceptionCounter + 1
        End If
    Next intColumnName
    GetArrayOfMissingColumns = arrColumnsWithExceptions
    Erase arrColumns
End Function

Private Function EnablePowerPivot() As Boolean
'Function needs to be fixed: in case "Power Pivot" add-in is not correctly loaded for some reason, the function still returns "True"

    Dim bAvailable As Boolean
    Dim comPowerPivot As COMAddIn
    Dim cmd As CommandBarControl
    
    On Error Resume Next
    Set cmd = Application.CommandBars("Worksheet Menu Bar").Controls(POWERPIVOT_MENUBAR_CONTROL)
    Err.Clear
    On Error GoTo 0
    
    If cmd Is Nothing Then
        Set comPowerPivot = Application.COMAddIns(POWERPIVOT_COMADDIN_PROGID)
        If Not comPowerPivot Is Nothing Then
            comPowerPivot.Connect = False
            If Not comPowerPivot.Connect Then comPowerPivot.Connect = True
            bAvailable = comPowerPivot.Connect
        End If
    End If
'"Assaf, when stop hits, means you need to add code to to refresh commadin"
    
    EnablePowerPivot = bAvailable
    Set comPowerPivot = Nothing
    Set cmd = Nothing
End Function

Private Sub MergeSameCells(WorkRange As Range)
    Dim cell As Range
    'turn off display alerts while merging
    Application.DisplayAlerts = False
    
    'merge all same cells in range
MergeSame:
    If WorkRange.Rows.count = 1 Then
        For Each cell In WorkRange
            If cell.Value = cell.Offset(0, 1).Value And Not IsEmpty(cell) Then
                Range(cell, cell.Offset(0, 1)).Merge
                cell.HorizontalAlignment = xlCenter
                GoTo MergeSame
            End If
        Next
    ElseIf WorkRange.Columns.count = 1 Then
        For Each cell In WorkRange
            If cell.Value = cell.Offset(1, 0).Value And Not IsEmpty(cell) Then
                Range(cell, cell.Offset(1, 0)).Merge
                cell.VerticalAlignment = xlVAlignCenter
                GoTo MergeSame
            End If
        Next
    End If
    
    'turn display alerts back on
    Application.DisplayAlerts = True
End Sub

Private Function CountFilesInFolder(folderPath As String, Optional FileExtension As String = "csv") As Integer
    Dim fileName As String
    Dim intFileCount As Integer
    
    ' Check if the folder path ends with a backslash, if not, add it
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Set the initial file count to 0
    intFileCount = 0
    
    ' Get the first file in the folder
    fileName = Dir(folderPath & "*." & FileExtension)
    
    ' Loop through all files in the folder
    Do While fileName <> ""
        ' Increment the file count
        intFileCount = intFileCount + 1
        
        ' Get the next file in the folder
        fileName = Dir()
    Loop
    
    CountFilesInFolder = intFileCount
End Function

Private Sub RenameCodeName(Sht As Worksheet, NewName As String)
    Dim VBProj As VBIDE.VBProject
    Dim vbComps As VBIDE.VBComponents
    Dim VBComp As VBIDE.VBComponent
    Dim vbProps As VBIDE.Properties
    Dim CodeNameProp As VBIDE.Property
    
    Set VBProj = Sht.Parent.VBProject
    Set vbComps = VBProj.VBComponents
    Set VBComp = vbComps(Sht.CodeName)
    Set vbProps = VBComp.Properties
    Set CodeNameProp = vbProps("_Codename")
    CodeNameProp.Value = NewName
    
    Set CodeNameProp = Nothing
    Set vbProps = Nothing
    Set VBComp = Nothing
    Set vbComps = Nothing
    Set VBProj = Nothing
End Sub

Private Sub DeleteIrrelevantRecords(Sht As Worksheet, FieldName As String, Criteria As Variant)
    Dim lngFieldColumnIndex As Long
    Dim lngEria As Long
    Dim erias As Range
    
    lngFieldColumnIndex = GetSheetColumnIndexByTitle(FieldName, Sht, Sht.Range("A1"))
    
    If lngFieldColumnIndex > 0 Then
        With Sht.Range("A1")
            .CurrentRegion.AutoFilter Field:=lngFieldColumnIndex, Criteria1:=Criteria, Operator:=xlFilterValues
            If AutoFilterRecordsFound(Sht) Then
                Set erias = Range(.Offset(1), Sht.Cells(.SpecialCells(xlCellTypeLastCell).Row, 1)).SpecialCells(xlCellTypeVisible)
                For lngEria = erias.Areas.count To 1 Step -1
                    erias.Areas(lngEria).EntireRow.Delete
                    'Range(.Offset(1), Sht.Cells(.SpecialCells(xlCellTypeLastCell).Row, 1)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
                Next lngEria
                'Range(.Offset(1), Sht.Cells(.SpecialCells(xlCellTypeLastCell).Row, 1)).SpecialCells(xlCellTypeVisible).EntireRow.Delete
            End If
            
            .Parent.ShowAllData
        End With
    End If
End Sub

Private Function AutoFilterRecordsFound(Sht As Worksheet) As Boolean
    Dim lngAreasFound As Long
    
    With Sht.AutoFilter.Range
        lngAreasFound = .SpecialCells(xlCellTypeVisible).Areas.count
        AutoFilterRecordsFound = lngAreasFound > 1 Or .SpecialCells(xlCellTypeVisible).Rows.count > 1
    End With
End Function

Sub RemovePivotTableSubtotals(pt As PivotTable)
    Dim pvtField As PivotField
    
    On Error Resume Next
    For Each pvtField In pt.PivotFields
        If pvtField.Orientation = xlColumnField Or pvtField.Orientation = xlRowField Then
            With pvtField
                .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                If LCase(pvtField.name) <> "year" And LCase(pvtField.name) <> "week" Then
                    .ShowAllItems = True
                End If
            End With
        End If
        Err.Clear
    Next pvtField
    On Error GoTo 0
End Sub

Sub RemoveDivisionByZeroColumns()
    Dim lngAreasCount As Long
    Dim lngColIndex As Long
    Dim rngData As Range
    
    Set rngData = Range("a1").CurrentRegion
    For lngColIndex = rngData.Columns.count To 1 Step -1
        On Error Resume Next 'ignore errors 'No cells were found'
        lngAreasCount = rngData.Columns(lngColIndex).SpecialCells(xlCellTypeConstants, 16).Areas.count
        Err.Clear
        On Error GoTo 0
        If lngAreasCount = 1 Then
            rngData.Columns(lngColIndex).EntireColumn.Delete
        End If
    Next lngColIndex
End Sub
