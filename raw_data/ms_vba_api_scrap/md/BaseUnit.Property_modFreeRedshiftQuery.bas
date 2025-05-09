Attribute VB_Name = "modFreeRedshiftQuery"
Option Explicit

Sub FreeRedshiftQuery()
Attribute FreeRedshiftQuery.VB_ProcData.VB_Invoke_Func = "R\n14"
    Dim qry As queryTable
    Dim wbkReport As Workbook
    Dim shtNew As Worksheet
    Dim strSqlQuery As String
'    Dim strSessionId As String
'    Dim strErrorDescription As String
    Dim Sht As Worksheet
    Dim strQueryName As String
    Dim strDSN As String
    
    strDSN = "Redshift_EU"
    strQueryName = ActiveCell.Offset(0, -1).Value2
    strSqlQuery = ActiveCell.Value

'    If IsSQLStatementValid(strSqlQuery, strErrorDescription) Then
'ExecuteQuery:
'        Application.ScreenUpdating = False
'
        Set wbkReport = Workbooks.Add(xlWBATWorksheet)
        Set shtNew = wbkReport.ActiveSheet
        Set qry = CreateQueryTable(shtNew, strDSN)
        With qry
            .CommandText = strSqlQuery
            .AdjustColumnWidth = True
            .Refresh BackgroundQuery:=False
        End With
        Call FormatDateColumns(shtNew)
        shtNew.name = Left(strQueryName, 31)
        
'**************************************
'        Call SplitWorksheetsByColumnValues("business", ActiveSheet)
'
'        For Each Sht In Worksheets
'            Call CreateChart(Sht)
'        Next Sht
'**************************************

        Application.ScreenUpdating = True
'    Else
'        If InStr(strErrorDescription, "cancelled on user's request") <> 0 Then
'            GoTo ExecuteQuery
'        Else
'            MsgBox strErrorDescription
'        End If
'    End If
End Sub

Function ReadTextFile(filePath As String) As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim fileLine As String
    
    ' Open the text file for reading
    fileNum = FreeFile
    Open filePath For Input As fileNum
    
    ' Read the content of the file
    Do Until EOF(fileNum)
        Line Input #fileNum, fileLine
        fileContent = fileContent & fileLine & vbCrLf
    Loop
    
    ' Close the file
    Close fileNum
    
    ' Remove the trailing newline (if any)
    If Right(fileContent, 2) = vbCrLf Then
        fileContent = Left(fileContent, Len(fileContent) - 2)
    End If
    
    ' Return the file content as a string
    ReadTextFile = fileContent
End Function

Public Sub CreateChart(Sht As Worksheet)
    Dim lngColIndexSessions As Long
    Dim rngDataNoHeaders As Range
    Dim Pnt As Point
    Dim lngPoint As Long
    Dim chartObj As ChartObject
    Dim MyShape As Shape
    Dim lngColumnTitleColIndex As Long
    Dim lngDateColIndex As Long
    Dim lngBaselineColIndex As Long
    Dim rngSourceData As Range
    Dim Serie As series
    Dim dblNumberSessions As Double
    Dim label As DataLabel
    Dim i As Long
    
    Set MyShape = Sht.Shapes.AddChart2(201, xlColumnClustered)
    
    Set rngSourceData = Application.Intersect(Sht.Columns("E:H"), Sht.UsedRange)

    With MyShape
        .LockAspectRatio = msoTrue
        .ScaleWidth 1.3, msoFalse
    End With
    
    With MyShape.Chart
        '.PlotVisibleOnly = False
        Set chartObj = .Parent
        .SetSourceData Source:=rngSourceData
        .PlotBy = xlColumns
        
        .ChartArea.Font.Size = 12
        
        For Each Serie In .FullSeriesCollection
            .ApplyDataLabels
        Next Serie
        
        Set Serie = .FullSeriesCollection(1)
        With Serie
            .ChartType = xlColumnStacked
            .AxisGroup = 1
            .Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
        End With
        With .FullSeriesCollection(2)
            .ChartType = xlColumnStacked
            .AxisGroup = 1
        End With
        With .FullSeriesCollection(3)
            .ChartType = xlLine
            .AxisGroup = 2
        End With
        .SetElement (msoElementChartTitleNone)
        
        With .Axes(xlValue, xlSecondary)
            .MinimumScale = 0
            .MaximumScale = 1.2
            .TickLabels.NumberFormat = "0%"
        End With
    
        With .FullSeriesCollection(3)
            .DataLabels.NumberFormat = "0%"
            For i = 1 To .DataLabels.count
                Set label = .DataLabels(i)
                label.Top = MyShape.Chart.PlotArea.Top + 5
            Next i
        End With
        
        .ChartGroups(1).GapWidth = 30
    
        With .FullSeriesCollection(2).DataLabels.Format.TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 12
        End With
    
        With .FullSeriesCollection(1).DataLabels.Format.TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 12
        End With

        With .FullSeriesCollection(3).DataLabels.Format.TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 12
        End With
        
        With .Legend.Format.TextFrame2.TextRange.Font
            .Size = 12
        End With
    
        With .Axes(xlCategory)
            .TickLabels.NumberFormat = "[$-fr-FR]mmm-yy;@"
            .BaseUnit = xlMonths
'            .Format.TextFrame2.TextRange.Font.Size = 12
        End With
    End With
    
    With chartObj
        .Left = 10 + (Sht.Shapes.count - 1) * (Sht.Shapes(1).Width + 10)
        .Top = .Top + 100
    End With
End Sub

