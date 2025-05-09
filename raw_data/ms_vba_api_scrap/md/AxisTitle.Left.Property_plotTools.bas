Attribute VB_Name = "plotTools"
'plotTools.bas (UDF)
'Collection of User Defined Functions (UDFs) and Subroutines for creating/configuring charts/graphs and plotting series data
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
Function getAxisBounds(maxVal As Double, ByVal roundToNearest As Double, getMax As Boolean) As Double
    If getMax Then
        getAxisBounds = round((maxVal + roundToNearest / 2) / roundToNearest) * roundToNearest
    Else
        getAxisBounds = round((maxVal - roundToNearest / 2) / roundToNearest) * roundToNearest
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function parseParameterString(parameterString As Variant, parameterIdentifier As String) As String
    Dim matchString1, matchString2, parsedString As String
    Dim i, startPos, matchPos1, matchPos2, remainingChars, lenMatchStrings As Integer
    matchString1 = parameterIdentifier + "="
    matchString2 = parameterIdentifier + "_"
    lenMatchStrings = Len(matchString1)
    
    If Not IsNumeric(parameterString) Then
        matchPos1 = InStr(parameterString, matchString1)
        matchPos2 = InStr(parameterString, matchString2)
        startPos = Application.WorksheetFunction.Max(matchPos1, matchPos2)
        If startPos > 0 Then
            parameterString = Mid(parameterString, startPos + lenMatchStrings, Len(parameterString))
        End If
        If Not IsNumeric(parameterString) Then
            remainingChars = Len(parameterString)
            For i = 1 To remainingChars
                parameterString = Left(parameterString, Len(parameterString) - 1)
                If IsNumeric(parameterString) Then
                    Exit For
                End If
            Next i
        End If
    End If
    parseParameterString = parameterString
End Function
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
Function createChart() As String
    ActiveSheet.Shapes.AddChart2(332, xlLineMarkers).Select
    Dim numCharts As Integer: numCharts = countWorkbookCharts()
    Dim chartHandle As Chart: Set chartHandle = ActiveChart
    createChart = "Chart " & Str(numCharts)
    ActiveChart.Parent.Name = createChart
End Function
Sub createTitledChart(plotTypeTitleLabel As String, iterations As Variant, materialDependency As String, titleSuffix As String, titleFontSize As Integer, suffixOnlyLine As Boolean, outputCell As Range)
    'Application.ScreenUpdating = False
    'Application.EnableEvents = False
    Dim iterationString
    Dim chartName As String: chartName = createChart()
    Dim chartHandle As Chart: Set chartHandle = ActiveChart
    'Selection.Top = Range(Application.Caller.Address).Top
    'Selection.Left = Range(Application.Caller.Address).Left
    Selection.Top = Range(outputCell.Address).Top
    Selection.Left = Range(outputCell.Address).Left
    iterationString = IIf(iterations = 1, iterations & " Iteration ", iterations & " Iterations ")
    'ActiveChart.Parent.Width = 382.24 '3.92 * ptsPerInch
    'ActiveChart.Parent.Height = 229.68 '3.19 * ptsPerInch
    'ActiveChart.ChartTitle.Text = plotTypeTitleLabel
    'ActiveChart.ChartTitle.Left = 0
    ActiveChart.ChartTitle.Text = plotTypeTitleLabel & " vs. N" & materialDependency & IIf(suffixOnlyLine, Chr(13) & "After " & iterationString & titleSuffix, " After " & iterationString & Chr(13) & titleSuffix)
    ActiveChart.ChartTitle.Font.Size = titleFontSize
End Sub
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
Function setSeries(xVals As Range, seriesVals As Range, seriesNum As Integer, seriesName As String, Optional addSeries As Boolean = True)
    If addSeries Then
        ActiveChart.SeriesCollection.NewSeries
    End If
    ActiveChart.FullSeriesCollection(seriesNum).Name = seriesName
    ActiveChart.FullSeriesCollection(seriesNum).XValues = xVals
    ActiveChart.FullSeriesCollection(seriesNum).Values = seriesVals
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function addOTVSseries(OTVS As Range, xVals As Range, seriesNum As Integer)
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(seriesNum).XValues = xVals
    If OTVS.Cells.count = 1 Then
        Dim i As Integer
        Dim OTVSarray As Variant
        ReDim OTVSarray(xVals.Cells.count - 1)
        For i = 0 To xVals.Cells.count - 1
            OTVSarray(i) = OTVS.Cells(1, 1).value
        Next i
        ActiveChart.FullSeriesCollection(seriesNum).Name = "OTVS"
        ActiveChart.FullSeriesCollection(seriesNum).Values = OTVSarray
    Else
        ActiveChart.FullSeriesCollection(seriesNum).Name = "OTVS"
        ActiveChart.FullSeriesCollection(seriesNum).Values = OTVS
    End If
    'Set subRange = r.Parent.Range(r.Cells(startPos), r.Cells(endPos))
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Sub addChartLegend()
    Application.Calculation = xlManual
    ActiveChart.HasLegend = True
    ActiveChart.Legend.Position = xlBottom 'xlLegendPositionRight
    With ActiveChart.Legend.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = 0
    End With
End Sub
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
Sub fillTVcheckSeries(xVals As Range, ser1 As Range, ser2 As Range, OTVS As Range)
    Call setSeries(xVals, ser1, 1, "NTVS, TV Check Off", False)
    Call setSeries(xVals, ser2, 2, "NTVS, TV Check On")
    Call addOTVSseries(OTVS, xVals, 3)
    '-------------------------------------------------------------------------------------
    Call setSeriesFormat(1, msoThemeColorBackground2, -0.5, msoLineSolid, 3, 7, msoThemeColorText1)
    Call setSeriesFormat(2, msoThemeColorBackground2, -0.1000000015, msoLineSolid, 1, 6, msoThemeColorBackground2)
    Call setSeriesFormat(3, msoThemeColorText1, 0, msoLineSysDash, -4142, 3, msoThemeColorText1)
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
Sub fillTVvsKSeries(xVals As Range, ser1 As Range, OTVS As Range)
    Call setSeries(xVals, ser1, 1, "NTVS, TV Check Off", False)
    Call addOTVSseries(OTVS, xVals, 2)
    '-------------------------------------------------------------------------------------
    Call setSeriesFormat(1, msoThemeColorBackground2, -0.15, msoLineSolid, 8, 7, msoThemeColorText1)
    Call setSeriesFormat(2, msoThemeColorText1, 0, msoLineSysDash, -4142, 3, msoThemeColorText1)
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
Sub fillAlphaSeries(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, ser4 As Range, ser5 As Range, OTVS As Range)
    Call setSeries(xVals, ser1, 1, ChrW(945) & "=0.5", False)
    Call setSeries(xVals, ser2, 2, ChrW(945) & "=0.65")
    Call setSeries(xVals, ser3, 3, ChrW(945) & "=0.75")
    Call setSeries(xVals, ser4, 4, ChrW(945) & "=0.85")
    Call setSeries(xVals, ser5, 5, ChrW(945) & "=0.95")
    Call addOTVSseries(OTVS, xVals, 6)
    '-------------------------------------------------------------------------------------
    Call setSeriesFormat(1, msoThemeColorText1, 0, msoLineSolid, 3, 7, msoThemeColorText1)
    Call setSeriesFormat(2, msoThemeColorBackground2, -0.15, msoLineSolid, 1, 6, msoThemeColorText1)
    Call setSeriesFormat(3, msoThemeColorBackground2, -0.35, msoLineSolid, 2, 7, msoThemeColorText1)
    Call setSeriesFormat(4, msoThemeColorBackground2, -0.15, msoLineSolid, 4, 7, msoThemeColorText1)
    Call setSeriesFormat(5, msoThemeColorBackground2, -0.5, msoLineSolid, 8, 7, msoThemeColorText1)
    Call setSeriesFormat(6, msoThemeColorText1, 0, msoLineSysDash, -4142, 3, msoThemeColorText1)
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
Sub fillvsLSeries(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range)
    Call setSeries(xVals, ser1, 1, ChrW(955) & "=0.0001", False)
    Call setSeries(xVals, ser2, 2, ChrW(955) & "=0.00015")
    Call setSeries(xVals, ser3, 3, ChrW(955) & "=0.0002")
    '-------------------------------------------------------------------------------------
    Call setSeriesFormat(1, msoThemeColorBackground2, -0.05, msoLineSolid, 3, 7, msoThemeColorText1)
    Call setSeriesFormat(2, msoThemeColorBackground2, -0.15, msoLineSolid, 8, 6, msoThemeColorText1)
    Call setSeriesFormat(3, msoThemeColorBackground2, -0.75, msoLineSolid, 1, 6, msoThemeColorText1)
End Sub
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
Sub setSeriesFormat(seriesNum As Integer, lineColor As MsoThemeColorIndex, brightness As Double, dashStyle As MsoLineDashStyle, markerStyle As Integer, markerSize As Integer, markerColor As MsoThemeColorIndex)
    With ActiveChart.FullSeriesCollection(seriesNum).Format.Line
        .Visible = msoFalse
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = lineColor
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = brightness
        .Transparency = 0
        .dashStyle = dashStyle
        .Weight = 2#
    End With
    With ActiveChart.FullSeriesCollection(seriesNum)
        .MarkerForegroundColor = markerColor
        .MarkerBackgroundColorIndex = 15
        .markerStyle = markerStyle
        .markerSize = markerSize
    End With
End Sub
Sub setAxesFormat(axisType As XlAxisType, axisTitle As String, fontSize As Integer, reverseValOrder As Boolean)
    ActiveChart.Axes(axisType).HasTitle = True
    ActiveChart.Axes(axisType, xlPrimary).axisTitle.Text = axisTitle
    ActiveChart.Axes(axisType).axisTitle.Font.Size = fontSize
    ActiveChart.Axes(axisType).HasMajorGridlines = True
    ActiveChart.Axes(axisType).MajorGridlines.Format.Line.Visible = msoTrue
    ActiveChart.Axes(axisType).MajorTickMark = xlOutside
    If axisType = xlValue Then
        ActiveChart.Axes(xlValue).HasMinorGridlines = True
        ActiveChart.Axes(xlValue).MinorGridlines.Format.Line.Visible = msoTrue
        ActiveChart.Axes(xlValue).MinorTickMark = xlOutside
        If reverseValOrder Then
            ActiveChart.Axes(xlValue).ReversePlotOrder = True
            ActiveChart.Axes(xlValue).Crosses = xlMaximum
        Else
            ActiveChart.Axes(xlValue).Crosses = xlMinimum
        End If
    Else
        ActiveChart.Axes(xlCategory).HasMinorGridlines = False
        ActiveChart.Axes(xlCategory).MinorGridlines.Format.Line.Visible = msoFalse
        ActiveChart.Axes(xlCategory).MinorTickMark = xlNone
        ActiveChart.Axes(xlCategory).AxisBetweenCategories = False
        If reverseValOrder Then
            ActiveChart.Axes(xlCategory).TickLabelPosition = xlHigh
        Else
            ActiveChart.Axes(xlCategory).TickLabelPosition = xlLow
        End If
    End If
    With ActiveChart.Axes(axisType).Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground2 'msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.brightness = -0.25 '0
    End With
End Sub
Sub setTVcheckChartSizing(titleLeftPos As Double)
    'ActiveChart.Parent.Width = 282.24 '3.92 * ptsPerInch
    ActiveChart.Parent.Width = 300#  '3.92 * ptsPerInch
    ActiveChart.Parent.Height = 229.68 '3.19 * ptsPerInch
    ActiveChart.PlotArea.Top = 34.18
    ActiveChart.PlotArea.Height = 146.846
    ActiveChart.PlotArea.Width = 269.846
    ActiveChart.PlotArea.Left = 83.18
    ActiveChart.Axes(xlCategory).axisTitle.Top = 180.224
    ActiveChart.Axes(xlValue).axisTitle.Left = 0
    ActiveChart.ChartTitle.Left = titleLeftPos
    ActiveChart.ChartTitle.Top = -1
    ActiveChart.Legend.Left = 0.989
    ActiveChart.Legend.Width = 288.26
    ActiveChart.Legend.Height = 20.846
    ActiveChart.Legend.Top = 203
End Sub
Sub setTVvsKChartSizing(titleLeftPos As Double)
    Call setTVcheckChartSizing(titleLeftPos)
    ActiveChart.ChartTitle.Left = titleLeftPos
    ActiveChart.ChartTitle.Top = -5
    ActiveChart.PlotArea.Top = 31.18
    ActiveChart.PlotArea.Height = 143.846
    ActiveChart.Axes(xlCategory).axisTitle.Top = 175.224
    ActiveChart.Legend.Width = 250.26
    ActiveChart.Legend.Left = 10.989
    ActiveChart.Legend.Height = 20.846
    ActiveChart.Legend.Top = 198
End Sub
Sub setLChartSizing(titleLeftPos As Double)
    ActiveChart.Parent.Width = 282.24 '3.92 * ptsPerInch
    ActiveChart.Parent.Height = 229.68 '3.19 * ptsPerInch
    ActiveChart.PlotArea.Top = 38.18
    ActiveChart.PlotArea.Height = 146.846
    ActiveChart.PlotArea.Width = 251.846
    ActiveChart.PlotArea.Left = 105.18
    ActiveChart.Axes(xlCategory).axisTitle.Top = 183.224
    ActiveChart.Axes(xlValue).axisTitle.Left = 0
    ActiveChart.ChartTitle.Left = titleLeftPos
    ActiveChart.ChartTitle.Top = 0
    ActiveChart.Legend.Left = 0.989
    ActiveChart.Legend.Width = 270.26
    ActiveChart.Legend.Height = 16.846
    ActiveChart.Legend.Top = 203
End Sub
Sub setAlphaChartSizing(titleLeftPos As Double)
    ActiveChart.Parent.Width = 370.8 '337.68 '5.15 * ptsPerInch
    ActiveChart.Parent.Height = 229.68 '258.48 '4.15 * ptsPerInch
    ActiveChart.PlotArea.Top = 41.68
    ActiveChart.PlotArea.Height = 146.846 '176.846
    ActiveChart.PlotArea.Width = 339.846 '309.846
    ActiveChart.PlotArea.Left = 105.18
    ActiveChart.Axes(xlCategory).axisTitle.Top = 188.224 '215.724
    ActiveChart.Axes(xlValue).axisTitle.Left = 0
    ActiveChart.ChartTitle.Left = titleLeftPos '85.467 '69.467
    ActiveChart.ChartTitle.Top = -5
    ActiveChart.Legend.Left = -3
    ActiveChart.Legend.Width = 367.26
    ActiveChart.Legend.Height = 19.846
    ActiveChart.Legend.Top = 272
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
Sub setChartSizing(plotType As String, plotTitle As String)
    Dim ChartSizingFunction As String, titleLeftPos As Double, titleLeftPosArray() As Variant
    '------------------------------------------------------------------------------------------------------------------------------------------
    If plotTitle = "Total Variation" Then
        titleLeftPosArray = Array(35.467, 35.467, 56.467, 50.467)
    ElseIf plotTitle = "RSP % Error" Then
        titleLeftPosArray = Array(35.467, 45.467, 61.467, 85.467)
    ElseIf plotTitle = "Standard Deviation" Then
        titleLeftPosArray = Array(57.467, 45.467, 55.467, 76.467)
    Else
        titleLeftPosArray = Array(35.467, 45.467, 45.467, 85.467)
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------
    If plotType = "TVcheck" Then
        ChartSizingFunction = "setTVcheckChartSizing"
        titleLeftPos = titleLeftPosArray(0)
    ElseIf plotType = "K" Then
        ChartSizingFunction = "setTVvsKChartSizing"
        titleLeftPos = titleLeftPosArray(1)
    ElseIf plotType = "L" Then
        ChartSizingFunction = "setLChartSizing"
        titleLeftPos = titleLeftPosArray(2)
    ElseIf plotType = "Alpha" Then
        ChartSizingFunction = "setAlphaChartSizing"
        titleLeftPos = titleLeftPosArray(3)
    Else
        ChartSizingFunction = "setTVcheckChartSizing"
        titleLeftPos = titleLeftPosArray(0)
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------
    Application.Run ChartSizingFunction, titleLeftPos
End Sub
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
Sub setPlotDataRanging(plotType As String, plotTitle As String, plotDataRange As Range)
    Dim roundingTo As Double, roundingToArray() As Variant
    '------------------------------------------------------------------------------------------------------------------------------------------
    If plotTitle = "Total Variation" Then
        roundingToArray = Array(50#, 50#, 200#, 500#)
    ElseIf plotTitle = "RSP % Error" Then
        roundingToArray = Array(0.1, 0.1, 0.1, 0.2)
    ElseIf plotTitle = "Standard Deviation" Then
        roundingToArray = Array(0.0005, 0.0005, 0.001, 0.001)
    Else
        roundingToArray = Array(0.1, 0.1, 0.1, 0.1)
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------
    If plotType = "TVcheck" Then
        roundingTo = roundingToArray(0)
    ElseIf plotType = "K" Then
        roundingTo = roundingToArray(1)
    ElseIf plotType = "L" Then
        roundingTo = roundingToArray(2)
    ElseIf plotType = "Alpha" Then
        roundingTo = roundingToArray(3)
    Else
        roundingTo = roundingToArray(0)
    End If
    '------------------------------------------------------------------------------------------------------------------------------------------
    Call queryPlotRange(plotTitle, plotDataRange, roundingTo)
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
Function queryPlotRange(plotTitle As String, plotDataRange As Range, roundingTo As Double)
    Dim minDataVal, maxDataVal, roundedMinDataVal, roundedMaxDataVal, majorAxisUnits, newMajorAxisUnits As Variant
    
    Call echoAssign(Application.WorksheetFunction.Min(plotDataRange), minDataVal, "minDataVal", False)
    Call echoAssign(Application.WorksheetFunction.Max(plotDataRange), maxDataVal, "maxDataVal", False)
    ActiveChart.Axes(xlValue).MinimumScale = minDataVal
    ActiveChart.Axes(xlValue).MaximumScale = maxDataVal
    Call echoAssign(ActiveChart.Axes(xlValue).MajorUnit, newMajorAxisUnits, "newMajorAxisUnits", False)
    
    Do While newMajorAxisUnits <> majorAxisUnits
        Call echoAssign(GetRoundedNumber(minDataVal - newMajorAxisUnits / 4, newMajorAxisUnits, RoundingDirection.Down), roundedMinDataVal, "roundedMinDataVal" & roundedMinDataVal, False)
        'Call echoAssign(GetRoundedNumber(minDataVal, newMajorAxisUnits, RoundingDirection.Down), roundedMinDataVal, "roundedMinDataVal" & roundedMinDataVal, False)
        Call echoAssign(GetRoundedNumber(maxDataVal + newMajorAxisUnits / 2, newMajorAxisUnits, RoundingDirection.Up), roundedMaxDataVal, "roundedMaxDataVal" & roundedMaxDataVal, False)
        
        'Call echoAssign(GetRoundedNumber(minDataVal - newMajorAxisUnits / 4, newMajorAxisUnits, RoundingDirection.Down), roundedMinDataVal, "roundedMinDataVal" & roundedMinDataVal, False)
        'Call echoAssign(GetRoundedNumber(maxDataVal + newMajorAxisUnits / 4, newMajorAxisUnits, RoundingDirection.Up), roundedMaxDataVal, "roundedMaxDataVal" & roundedMaxDataVal, False)
        'Call echoAssign(GetRoundedNumber(maxDataVal + newMajorAxisUnits / 2 + 0.00000001, newMajorAxisUnits, RoundingDirection.Nearest), roundedMaxDataVal, "roundedMaxDataVal" & roundedMaxDataVal, False)
        'Call echoAssign(GetRoundedNumber(maxDataVal + 0.00000001, newMajorAxisUnits, RoundingDirection.Up), roundedMaxDataVal, "roundedMaxDataVal" & roundedMaxDataVal, False)
        ActiveChart.Axes(xlValue).MinimumScale = roundedMinDataVal
        ActiveChart.Axes(xlValue).MaximumScale = roundedMaxDataVal
        Call echoAssign(newMajorAxisUnits, majorAxisUnits, "majorAxisUnits", False)
        Call echoAssign(ActiveChart.Axes(xlValue).MajorUnit, newMajorAxisUnits, "newMajorAxisUnits", False)
    Loop
    
    'Exit Function
End Function
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
Function popIterationsVal(seriesRange As Range, xVals As Range) As Range
    Set popIterationsVal = subRange(seriesRange, 1 + seriesRange.Cells.count - xVals.Cells.count, seriesRange.Cells.count)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function queryIterations(seriesRange As Range, xVals As Range, iterations As Variant) As Variant
    If seriesRange.Cells.count = 2 Or seriesRange.Cells.count > xVals.Cells.count Then
        queryIterations = seriesRange.Cells(1, 1).value
    Else
        queryIterations = iterations
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function querySeriesIterations(xVals As Range, ser1 As Range, Optional ser2 As Range, Optional ser3 As Range, Optional ser4 As Range, Optional ser5 As Range, Optional ser6 As Range) As Variant
    querySeriesIterations = -1
    querySeriesIterations = queryIterations(ser1, xVals, querySeriesIterations)
    If Not IsMissing(ser2) Then
        querySeriesIterations = queryIterations(ser2, xVals, querySeriesIterations)
    ElseIf Not IsMissing(ser3) Then
        querySeriesIterations = queryIterations(ser3, xVals, querySeriesIterations)
    ElseIf Not IsMissing(ser4) Then
        querySeriesIterations = queryIterations(ser4, xVals, querySeriesIterations)
    ElseIf Not IsMissing(ser5) Then
        querySeriesIterations = queryIterations(ser5, xVals, querySeriesIterations)
    ElseIf Not IsMissing(ser6) Then
        querySeriesIterations = queryIterations(ser6, xVals, querySeriesIterations)
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function parseOTVSdata(OTVS As Range, xVals As Range) As Range
    If OTVS.Cells.count > xVals.Cells.count Then
        Set OTVS = subRange(OTVS, 1 + OTVS.Cells.count - xVals.Cells.count, OTVS.Cells.count)
    ElseIf OTVS.Cells.count = 2 Then
        Set OTVS = subRange(OTVS, 2, OTVS.Cells.count)
        'Set OTVS = OTVS.Cells(2, 1)
    End If
    Set parseOTVSdata = OTVS
End Function
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
Function configVsTVCheck(xVals As Range, ser1 As Range, ser2 As Range, OTVS As Range, Optional iterations As Variant) As Variant
    configVsTVCheck = IIf(IsMissing(iterations), querySeriesIterations(xVals, ser1, ser2, OTVS), iterations)
    Set OTVS = parseOTVSdata(OTVS, xVals)
    Set ser1 = popIterationsVal(ser1, xVals)
    Set ser2 = popIterationsVal(ser2, xVals)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function setVsTVCheckTitleParamString(paramValCells As Range) As Variant
    setVsTVCheckTitleParamString = ChrW(945) & "=" & parseParameterString(paramValCells.Item(1), "A") & ", "
    setVsTVCheckTitleParamString = setVsTVCheckTitleParamString & ChrW(955) & "=" & parseParameterString(paramValCells.Item(1), "L")
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Sub CreateVsTVCheckChart(plotTypeTitleLabel As String, plotTypeAxisLabel As String, xVals As Range, ser1 As Range, ser2 As Range, OTVS As Range, iterations As Variant, paramValCells As Range, material As Variant, outputCell As Range)
    Dim paramString As Variant: paramString = setVsTVCheckTitleParamString(paramValCells)
    Dim iterationString: iterationString = IIf(iterations = 1, "After " & iterations & " Iteration ", "After " & iterations & " Iterations ")
    If plotTypeTitleLabel = "Total Variation" Then
        Call createTitledChart(plotTypeTitleLabel, iterations, "", "(" & paramString & ")", 14, False, outputCell)
    Else
        Call createTitledChart(plotTypeTitleLabel, iterations, " in " & material, "(" & paramString & ")", 14, True, outputCell)
        'ActiveChart.ChartTitle.Text = plotTypeTitleLabel & " vs. N" & materialDependency & IIf(suffixOnlyLine, Chr(13) & "After " & iterationString & titleSuffix, " After " & iterationString & Chr(13) & titleSuffix)
        'ActiveChart.ChartTitle.Text = plotTypeTitleLabel & " vs. N " & " in "
    End If
    Call fillTVcheckSeries(xVals, ser1, ser2, OTVS)
    Call setAxesFormat(xlValue, plotTypeAxisLabel, 12, False)
    Call setAxesFormat(xlCategory, "N", 12, False)
    Call addChartLegend
    Call setChartSizing("TVcheck", plotTypeTitleLabel)
    '    ActiveChart.ChartTitle.Text = plotTypeTitleLabel & " vs. N " & " in " & material & Chr(13) & iterationString & "(" & paramString & ")"
    Dim plotDataRange As Range: Set plotDataRange = Union(ser1, ser2, OTVS)
    Call setPlotDataRanging("TVcheck", plotTypeTitleLabel, plotDataRange)
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotTVvsTVCheck(xVals As Range, ser1 As Range, ser2 As Range, OTVS As Range, paramValCells As Range, outputCell As Range, Optional iterations As Variant)
    iterations = configVsTVCheck(xVals, ser1, ser2, OTVS)
    Call CreateVsTVCheckChart("Total Variation", "Total Variation (TV)", xVals, ser1, ser2, OTVS, iterations, paramValCells, "", outputCell)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotRSPerrorVsTVCheck(xVals As Range, ser1 As Range, ser2 As Range, OTVS As Range, paramValCells As Range, material As Variant, outputCell As Range, Optional iterations As Variant)
    iterations = configVsTVCheck(xVals, ser1, ser2, OTVS)
    Call CreateVsTVCheckChart("RSP % Error", "RSP Error (%)", xVals, ser1, ser2, OTVS, iterations, paramValCells, material, outputCell)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotStdDevVsTVCheck(xVals As Range, ser1 As Range, ser2 As Range, OTVS As Range, paramValCells As Range, material As Variant, outputCell As Range, Optional iterations As Variant)
    iterations = configVsTVCheck(xVals, ser1, ser2, OTVS)
    Call CreateVsTVCheckChart("Standard Deviation", "Standard Deviation", xVals, ser1, ser2, OTVS, iterations, paramValCells, material, outputCell)
End Function
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
Function configVsK(xVals As Range, ser1 As Range, OTVS As Range, Optional iterations As Variant) As Variant
    configVsK = IIf(IsMissing(iterations), querySeriesIterations(xVals, ser1, OTVS), iterations)
    Set OTVS = parseOTVSdata(OTVS, xVals)
    Set ser1 = popIterationsVal(ser1, xVals)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function createTVvsKTable(plotTypeTitleLabel As String, plotTypeAxisLabel As String, xVals As Range, ser1 As Range, OTVS As Range, iterations As Variant, lambda As Variant, outputCell As Range)
    lambda = parseParameterString(lambda, "L")
    Call createTitledChart(plotTypeTitleLabel, iterations, "", ChrW(955) & "=" & lambda, 14, False, outputCell)
    Call fillTVvsKSeries(xVals, ser1, OTVS)
    Call setAxesFormat(xlValue, plotTypeAxisLabel, 12, False)
    Call setAxesFormat(xlCategory, "N", 12, False)
    Call addChartLegend
    Call setChartSizing("K", plotTypeTitleLabel)
    Dim plotDataRange As Range: Set plotDataRange = Union(ser1, OTVS)
    Call setPlotDataRanging("K", plotTypeTitleLabel, plotDataRange)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotTVvsK(xVals As Range, ser1 As Range, OTVS As Range, lambda As Variant, outputCell As Range, Optional iterations As Variant)
    iterations = configVsK(xVals, ser1, OTVS)
    Call createTVvsKTable("Total Variation", "Total Variation (TV)", xVals, ser1, OTVS, iterations, lambda, outputCell)
End Function
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
Function configVsL(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, Optional iterations As Variant) As Variant
    configVsL = IIf(IsMissing(iterations), querySeriesIterations(xVals, ser1, ser2, ser3), iterations)
    Set ser1 = popIterationsVal(ser1, xVals)
    Set ser2 = popIterationsVal(ser2, xVals)
    Set ser3 = popIterationsVal(ser3, xVals)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function setVsLTitleParamString(lambdaVals As Range) As Variant
    setVsLTitleParamString = ChrW(945) & "=" & parseParameterString(lambdaVals.Item(1), "A") & ", "
    setVsLTitleParamString = setVsLTitleParamString & ChrW(955) & "=" & parseParameterString(lambdaVals.Item(1), "L")
    setVsLTitleParamString = setVsLTitleParamString & "," & parseParameterString(lambdaVals.Item(2), "L")
    setVsLTitleParamString = setVsLTitleParamString & "," & parseParameterString(lambdaVals.Item(3), "L")
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function createVsLChart(plotTypeTitleLabel As String, plotTypeAxisLabel As String, xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, iterations As Variant, paramString As Variant, material As Variant, reverseValOrder As Boolean, outputCell As Range)
    If plotTypeTitleLabel = "Total Variation" Then
        Call createTitledChart(plotTypeTitleLabel, iterations, "", "(" & paramString & ")", 14, False, outputCell)
        ActiveChart.ChartTitle.Text = plotTypeTitleLabel & " vs. N " & Chr(13) & "(" & paramString & ")"
    Else
        Call createTitledChart(plotTypeTitleLabel, iterations, "in " & material, "(" & paramString & ")", 14, True, outputCell)
        ActiveChart.ChartTitle.Text = plotTypeTitleLabel & " vs. N in " & material & Chr(13) & "(" & paramString & ")"
    End If
    Call fillvsLSeries(xVals, ser1, ser2, ser3)
    Call setAxesFormat(xlValue, plotTypeAxisLabel, 12, False)
    Call setAxesFormat(xlCategory, "N", 12, False)
    Call addChartLegend
    Call setChartSizing("L", plotTypeTitleLabel)
    Dim plotDataRange As Range: Set plotDataRange = Union(ser1, ser2, ser3)
    Call setPlotDataRanging("L", plotTypeTitleLabel, plotDataRange)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotTVVsL(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, lambdaVals As Range, outputCell As Range, Optional iterations As Variant)
    Dim lambdaValsString As Variant: lambdaValsString = setVsLTitleParamString(lambdaVals)
    iterations = configVsL(xVals, ser1, ser2, ser3, iterations)
    Call createVsLChart("Total Variation", "Total Variation (TV)", xVals, ser1, ser2, ser3, iterations, lambdaValsString, "", False, outputCell)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Public Function plotStdDevVsL(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, lambdaVals As Range, material As Variant, outputCell As Range, Optional iterations As Variant)
    Dim lambdaValsString As Variant: lambdaValsString = setVsLTitleParamString(lambdaVals)
    iterations = configVsL(xVals, ser1, ser2, ser3, iterations)
    Call createVsLChart("Standard Deviation", "Standard Deviation", xVals, ser1, ser2, ser3, iterations, lambdaValsString, material, False, outputCell)
End Function
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
Sub createVsAlphaChart(plotTypeTitleLabel As String, plotTypeAxisLabel As String, xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, ser4 As Range, ser5 As Range, OTVS As Range, iterations As Variant, lambda As Variant, material As Variant, reverseValOrder As Boolean, outputCell As Range)
    lambda = parseParameterString(lambda, "L")
    If plotTypeTitleLabel = "Total Variation" Then
        Call createTitledChart(plotTypeTitleLabel, iterations, "", "(" & ChrW(955) & "=" & lambda & ")", 18, False, outputCell)
    Else
        Call createTitledChart(plotTypeTitleLabel, iterations, " in " & material, "(" & ChrW(955) & "=" & lambda & ")", 18, True, outputCell)
    End If
    Call fillAlphaSeries(xVals, ser1, ser2, ser3, ser4, ser5, OTVS)
    Call setAxesFormat(xlValue, plotTypeAxisLabel, 13, reverseValOrder)
    Call setAxesFormat(xlCategory, "N", 13, reverseValOrder)
    Call addChartLegend
    Call setChartSizing("Alpha", plotTypeTitleLabel)
    Dim plotDataRange As Range: Set plotDataRange = Union(ser1, ser2, ser3, ser4, ser5, OTVS)
    Call setPlotDataRanging("Alpha", plotTypeTitleLabel, plotDataRange)
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
Function configVsAlpha(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, ser4 As Range, ser5 As Range, OTVS As Range, Optional iterations As Variant) As Variant
    configVsAlpha = IIf(IsMissing(iterations), querySeriesIterations(xVals, ser1, ser2, ser3, ser4, ser5, OTVS), iterations)
    Set OTVS = parseOTVSdata(OTVS, xVals)
    Set ser1 = popIterationsVal(ser1, xVals)
    Set ser2 = popIterationsVal(ser2, xVals)
    Set ser3 = popIterationsVal(ser3, xVals)
    Set ser4 = popIterationsVal(ser4, xVals)
    Set ser5 = popIterationsVal(ser5, xVals)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotTVVsAlpha(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, ser4 As Range, ser5 As Range, OTVS As Range, lambda As Variant, outputCell As Range, Optional iterations As Variant)
    iterations = configVsAlpha(xVals, ser1, ser2, ser3, ser4, ser5, OTVS, iterations)
    Call createVsAlphaChart("Total Variation", "Total Variation (TV)", xVals, ser1, ser2, ser3, ser4, ser5, OTVS, iterations, lambda, "", False, outputCell)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotRSPerrorVsAlpha(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, ser4 As Range, ser5 As Range, OTVS As Range, lambda As Variant, material As String, reverseValOrder As Boolean, outputCell As Range, Optional iterations As Variant)
    'lambda = parseParameterString(lambda, "L")
    iterations = configVsAlpha(xVals, ser1, ser2, ser3, ser4, ser5, OTVS, iterations)
    Call createVsAlphaChart("RSP % Error", "RSP Error (%)", xVals, ser1, ser2, ser3, ser4, ser5, OTVS, iterations, lambda, material, reverseValOrder, outputCell)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
Function plotStdDevVsAlpha(xVals As Range, ser1 As Range, ser2 As Range, ser3 As Range, ser4 As Range, ser5 As Range, OTVS As Range, lambda As Variant, material As String, outputCell As Range, Optional iterations As Variant)
    'lambda = parseParameterString(lambda, "L")
    iterations = configVsAlpha(xVals, ser1, ser2, ser3, ser4, ser5, OTVS, iterations)
    Call createVsAlphaChart("Standard Deviation", "Standard Deviation", xVals, ser1, ser2, ser3, ser4, ser5, OTVS, iterations, lambda, material, False, outputCell)
End Function
'******************************************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************************************





