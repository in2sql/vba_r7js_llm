Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
    Private Declare PtrSafe Function CreateThread Lib "kernel32" ( _
        ByVal lpThreadAttributes As LongPtr, _
        ByVal dwStackSize As Long, _
        ByVal lpStartAddress As LongPtr, _
        ByVal lpParameter As LongPtr, _
        ByVal dwCreationFlags As Long, _
        ByRef lpThreadId As Long) As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Const ConfigTitleRangeLocation_str As String = "配!$1:$1"
Public Const ConfigTitleKeywordLocation_str As String = "配!$A:$A"
Public Const ConfigSheetName As String = "配"
Public Const RAMSheetName As String = "寄"
Public Const ConfigTitleKeyword_str As String = "明细!中代表【土地用途考核项】的字段名称是"
Public Const WORKBOOK_NAME_LOCATION_str As String = "汇总!$B$2"
Public Const ChartBuilderSheetName_str As String = "中P"
Public Const CitySummaryTableLocation_str As String = "中S!$M$5"
Public Const SummarySheetName_str As String = "汇总"
Public Const DetailSheetName_str As String = "明细"
Public Const KeyFiledsTitleLocation_str As String = "土地用途考核项"
Public Const KeyFiledsTitleLocationShort_str As String = "考核sheet简称"
Public Const ChartBuliderSourceLocation_str As String = "图源绝对位置"
Public Const SummaryTableSourceLocation_str As String = "汇总图源绝对"
Public Const SummaryTableSourceColumnCounts_str As String = "汇总图源列宽"
Public Const SummaryTable2Location_str As String = "汇总粘贴至绝对"
Public Const SummaryChartNameLocation_str As String = "汇总统计图名称"
Public Const SummaryChart2Location_str As String = "汇总统计图粘贴至"
Public Const OverviewKeywords_str As String = "概述："
Public Const SummarySheetKeyFieldTitleLocation_str As String = "汇总!$C$1"

Function Unique2RAMSheet() As Integer
    Dim fromRange As Range, toRange As Range, detailTitle As Range, RAMTitle As Range, columnName As String
    Set RAMTitle = Sheets(RAMSheetName).Range("$6:$6")
    Sheets(RAMSheetName).Range("$7:$1048576").ClearContents
    Dim cell_i As Range
    For Each cell_i In RAMTitle
        If cell_i.Value <> "" Then
            columnName = cell_i.Value
            Set detailTitle = Sheets(DetailSheetName_str).Range("$1:$1")
            Set fromRange = detailTitle.Find(columnName)
            Set fromRange = fromRange.Resize(1024000, 1)
            Dim ws As Worksheet
            Set ws = detailTitle.Parent
            ws.Activate
            fromRange.Copy
            Set toRange = RAMTitle.Find(columnName, LookIn:=xlValues)
            toRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            Set toRange = toRange.Resize(1024100, 1)
            toRange.RemoveDuplicates Columns:=1, Header:=xlYes
        Else
            Unique2RAMSheet = 0
            Exit Function
        End If
    Next cell_i
End Function

Function WithoutStarString() As Integer
    Sheets(DetailSheetName_str).Cells.Replace What:="~*", Replacement:="×", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    WithoutStarString = 0
End Function

Function GetConfigKeywordString(keywordLocation As String) As String
    Dim configKeywordRange As Range
    Dim row_index As Long
    Set configKeywordRange = Range(ConfigTitleKeywordLocation_str)
    row_index = 1
    Dim cell_i As Range
    For Each cell_i In configKeywordRange
        If CStr(cell_i.Value) = keywordLocation Then
            GetConfigKeywordString = CStr(configKeywordRange.Cells(row_index, 1).Offset(0, 1).Value)
            Exit Function
        Else
            row_index = row_index + 1
        End If
        If row_index > 65536 Then
            On Error GoTo ErrorNA
        End If
    Next cell_i
ErrorNA:
    MsgBox keywordLocation & Chr(10) & " -> Function@GetConfigKeywordString", , "结果不存在"
    On Error GoTo 0
End Function

Function GetConfigTitleRange(titleName As String) As Range
    Dim configTitleRange As Range
    Dim col_index As Long
    Set configTitleRange = Range(ConfigTitleRangeLocation_str)
    col_index = 1
    Dim cell_i As Range
    For Each cell_i In configTitleRange
        If CStr(cell_i.Value) = titleName Then
            Set GetConfigTitleRange = configTitleRange.Cells(1, col_index)
            Exit Function
        Else
            col_index = col_index + 1
        End If
        If col_index > 16384 Then
            On Error GoTo ErrorNA
        End If
    Next cell_i
ErrorNA:
    MsgBox titleName & Chr(10) & " -> Function@GetConfigTitleRange", , "结果不存在"
    On Error GoTo 0
End Function

Function RangeDown2BlankFast(firstCell As Range) As Range ' 不适合拓选包含公式的单元格
    Dim targetCell As Range
    Set targetCell = firstCell
    
    Dim lastCell As Range
    Set lastCell = firstCell.End(xlDown)
    
    If lastCell.Row > firstCell.Row Then
        Set RangeDown2BlankFast = firstCell.Resize(lastCell.Row - firstCell.Row + 1)
    Else
        Set RangeDown2BlankFast = firstCell
    End If
End Function

Function RangeDown2Blank(firstCell As Range) As Range
    Dim cellsCount As Long
    cellsCount = 0
    
    Dim targetCell As Range
    Set targetCell = firstCell
    
    Dim dataValues As Variant
    dataValues = firstCell.Resize(1048576 - firstCell.Row + 1).Value
    While CStr(dataValues(cellsCount + 1, 1)) <> ""
        cellsCount = cellsCount + 1
    Wend
    
    cellsCount = cellsCount - 1
    
    Set RangeDown2Blank = Range(firstCell.Worksheet.Name & "!" & firstCell.Address & ":" & firstCell.Offset(cellsCount, 0).Address)
End Function

Function RangeRight2Blank(rangeA As Range) As Range
    Dim lastCell As Range
    Set lastCell = rangeA.End(xlToRight)
    Set RangeRight2Blank = Range(rangeA.Address & ":" & lastCell.Address)
End Function

Function RangeExpand2Up(sourceRange As Range, targetRows As Long) As Range
    targetRows = -1 * targetRows
    sourceRange = sourceRange.Offset(targetRows, 0)
    sourceRange = sourceRange.Resize(sourceRange.Rows.Count + targetRows, sourceRange.Columns.Count)
    Set RangeExpand2Up = sourceRange
    Exit Function
End Function

Function CopyAndPasteAsValue(targetCell As Range) As Integer
    Dim sourceCell As Range, destinationCell As Range
    If targetCell Is Nothing Then
        Exit Function
    End If
    Set sourceCell = targetCell.Offset(-1, 1)
    Set destinationCell = targetCell
    Dim ws As Worksheet
    Set ws = sourceCell.Parent
    ws.Activate
    sourceCell.Copy
    destinationCell.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Function

Function FindDownByRowWithStartText(firstCell As Range, startText As String) As Range
'不会返回自身；间隔超过1024会错误
    Dim dataArr As Variant
    dataArr = firstCell.Resize(1024, 1).Value
    
    Dim i As Long
    For i = 1 To UBound(dataArr, 1)
        Dim currentCellStr As String
        currentCellStr = CStr(dataArr(i, 1))
        
        If Left(currentCellStr, Len(startText)) = startText Then
            Set FindDownByRowWithStartText = firstCell.Offset(i - 1, 0)
            Exit Function
        End If
    Next i
End Function

Function GetStartEndIndexOfText(searchStr As String, text As String) As Variant
    Dim startPos As Integer
    Dim endPos As Integer
    Dim result(1 To 2) As Integer
    startPos = InStr(text, searchStr)
    endPos = startPos + Len(searchStr) - 1
    result(1) = startPos
    result(2) = endPos
    GetStartEndIndexOfText = result
    
End Function

Function FormatText_GS(targetCell As Range) As Integer
    Dim tempText As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim errorFlag As Boolean
    Dim searchTexts(1 To 5) As String, i As Integer
    errorFlag = False
    On Error GoTo ErrorText
        tempText = targetCell.Value
    
    If errorFlag = True Then
        Exit Function
    End If
    
    searchTexts(1) = "    整体"
    searchTexts(2) = "万平米，其中"
    searchTexts(3) = "家土地使用人占到了整体"
    searchTexts(4) = "其中拿地面积最大的是"
    searchTexts(5) = "，面积达到"
    
    targetCell.Font.Color = RGB(255, 0, 0)
    targetCell.Characters(1, 25).Font.Color = RGB(0, 0, 0)
    For i = 1 To (UBound(searchTexts) - LBound(searchTexts) + 1)
        startPos = InStr(tempText, searchTexts(i))
        If startPos > 0 Then
            endPos = startPos + Len(searchTexts(i)) - 1
            targetCell.Characters(startPos, Len(searchTexts(i))).Font.Color = RGB(0, 0, 0)
        End If
    Next i
    FormatText_GS = 0
ErrorText:
    errorFlag = True
    On Error GoTo 0
End Function

Function SplitSeries(seriesText As String) As Variant
    Dim startIndex As Long
    Dim endIndex As Long
    Dim rangeStringList(1 To 3) As String
    Dim withoutBracketText As String
    withoutBracketText = Mid(seriesText, 9, Len(seriesText) - 9)

    startIndex = InStr(withoutBracketText, ",")
    endIndex = InStr(startIndex + 1, withoutBracketText, ",")
    rangeStringList(1) = Mid(withoutBracketText, 1, startIndex - 1)
    rangeStringList(2) = Mid(withoutBracketText, startIndex + 1, endIndex - startIndex - 1)
    startIndex = InStr(endIndex + 1, withoutBracketText, ",")
    rangeStringList(3) = Mid(withoutBracketText, endIndex + 1, startIndex - endIndex - 1)
    
    SplitSeries = rangeStringList

End Function

Function ResetDataSourceAuto(chartName As String) As Integer
    Dim t
    t = Timer
    Dim ws As Worksheet
    Dim chtOBJ As ChartObject
    Dim dataSource As Range, dataSeries As String
    Dim xDataSource As Range
    Set ws = ThisWorkbook.Sheets(ChartBuilderSheetName_str)
    ws.Activate
    For Each chtOBJ In ws.ChartObjects
        If chtOBJ.Name = chartName Then
            dataSeries = chtOBJ.Chart.SeriesCollection(1).Formula
            Set dataSource = Range(SplitSeries(dataSeries)(3))(1, 1)
            Set dataSource = RangeDown2Blank(dataSource)
            Set xDataSource = Range(SplitSeries(dataSeries)(2))(1, 1)
            Set xDataSource = RangeDown2Blank(xDataSource)
            chtOBJ.Chart.SetSourceData dataSource
            chtOBJ.Chart.FullSeriesCollection(1).Name = "=" & SplitSeries(dataSeries)(1)
            chtOBJ.Chart.FullSeriesCollection(1).XValues = "=" & xDataSource.Worksheet.Name & "!" & xDataSource.Address
            ResetDataSourceAuto = 0
            Debug.Print "》》" & chartName & "》》数据源更新完成>==>耗时" & CStr(Timer - t)
            Exit Function
        End If
    Next chtOBJ
    MsgBox "未找到名称为'" & chartName & "'的图表"
End Function

Function 城市分布生成图专用函数(chartName As String) As Integer
    Dim t
    t = Timer
    Dim ws As Worksheet
    Dim chtOBJ As ChartObject
    Dim dataSource As Range, dataSeries As String
    Dim xDataSource As Range, newDataSource As Range
    Set ws = ThisWorkbook.Sheets(ChartBuilderSheetName_str)
    ws.Activate
    For Each chtOBJ In ws.ChartObjects
        If chtOBJ.Name = chartName Then
            dataSeries = chtOBJ.Chart.SeriesCollection(1).Formula
            Set dataSource = Range(SplitSeries(dataSeries)(3))(1, 1)
            Set dataSource = RangeDown2Blank(dataSource)
            Set xDataSource = Range(SplitSeries(dataSeries)(2))(1, 1)
            Set xDataSource = RangeDown2Blank(xDataSource)
            Set newDataSource = RangeRight2Blank(RangeDown2Blank(Range(CitySummaryTableLocation_str)))
            chtOBJ.Chart.SetSourceData Source:=Sheets(Range(CitySummaryTableLocation_str).Worksheet.Name).Range(newDataSource.Address)
            chtOBJ.Chart.PlotBy = xlColumns
            If newDataSource.Rows.Count > 12 Then
                chtOBJ.Chart.SetElement (msoElementDataLabelNone)
            Else
                chtOBJ.Chart.SetElement (msoElementDataLabelOutSideEnd)
            End If
            If ThisWorkbook.Sheets(ConfigSheetName).Range("$B$1") = "省级" Then
                chtOBJ.Chart.ChartTitle.text = "各城市属性分布"
            Else
                chtOBJ.Chart.ChartTitle.text = "各区县属性分布"
            End If
            Dim series_i As Series
            For Each series_i In chtOBJ.Chart.FullSeriesCollection
                If series_i.Name = "合计" Then
                    series_i.IsFiltered = True
                    Exit For
                End If
            Next series_i
            城市分布生成图专用函数 = 0
            Debug.Print "》》" & chartName & "》》数据源更新完成>==>耗时" & CStr(Timer - t)
            Exit Function
        End If
    Next chtOBJ
    MsgBox "未找到名称为'" & chartName & "'的图表"
End Function

Function AutoPastePicture1() As Integer
    '将表格粘贴为图片
    Dim t
    Dim tablePicSourceLocationRange As Range, tablePicPasteLocationRange As Range, tableSourceColumnCountRange As Range
    Set tablePicSourceLocationRange = RangeDown2Blank(GetConfigTitleRange(SummaryTableSourceLocation_str).Offset(1, 0))
    Set tablePicPasteLocationRange = RangeDown2Blank(GetConfigTitleRange(SummaryTable2Location_str).Offset(1, 0))
    Set tableSourceColumnCountRange = RangeDown2Blank(GetConfigTitleRange(SummaryTableSourceColumnCounts_str).Offset(1, 0))
    Dim tablePicSourceLocationList(1 To 32) As Variant
    Dim tablePicPasteLocationList(1 To 32) As Variant
    Dim tablePicColumnCountList(1 To 32) As Variant
    Dim i As Long
    For i = 1 To 32
        If CStr(tablePicSourceLocationRange.Cells(i, 1).Value) <> "" Then
            tablePicSourceLocationList(i) = tablePicSourceLocationRange.Cells(i, 1).Value
            tablePicPasteLocationList(i) = tablePicPasteLocationRange.Cells(i, 1).Value
            tablePicColumnCountList(i) = tableSourceColumnCountRange.Cells(i, 1).Value
        End If
    Next i
    
    For i = 1 To 32
        If CStr(tablePicSourceLocationList(i)) <> "" Then
            t = Timer
            Dim tablePicSourceRange As Range
            Dim tablePic As Picture, debug1 As String
            Set tablePicSourceRange = RangeDown2Blank(Range(CStr(tablePicSourceLocationList(i))))
            Set tablePicSourceRange = tablePicSourceRange.Resize(tablePicSourceRange.Rows.Count, tablePicColumnCountList(i))
            Dim tws As Worksheet
            Set tws = tablePicSourceRange.Parent
            tws.Activate
            Call TryCopyOBJ(tablePicSourceRange)
            Set tablePic = TryPicPaste2Sheet(ThisWorkbook.Sheets(SummarySheetName_str))
            With tablePic
                .Left = ThisWorkbook.Sheets(SummarySheetName_str).Range(tablePicPasteLocationList(i)).Left
                .Top = ThisWorkbook.Sheets(SummarySheetName_str).Range(tablePicPasteLocationList(i)).Top
            End With
            Dim targetHeight As Integer
            targetHeight = 300
            If tablePic.Height > targetHeight Then
                Dim scaleFactor As Double
                scaleFactor = targetHeight / tablePic.Height
                With tablePic
                    .Height = targetHeight
                    .Width = tablePic.Width * scaleFactor
                End With
            End If
            Dim rubbish3 As Integer
            ' rubbish3 = ClearWindowsClipboard()
            Debug.Print "》》" & CStr(tablePicPasteLocationList(i)) & "》》表格图粘贴完成>==>耗时" & CStr(Timer - t)
        End If
    Next i
    AutoPastePicture1 = 0
End Function

Function AutoPastePicture2() As Integer
    Dim t
    Dim chartPicNameLocationRange As Range, chartPicPasteLocationRange As Range, chartSheet As Worksheet
    Set chartPicNameLocationRange = RangeDown2Blank(GetConfigTitleRange(SummaryChartNameLocation_str).Offset(1, 0))
    Set chartPicPasteLocationRange = RangeDown2Blank(GetConfigTitleRange(SummaryChart2Location_str).Offset(1, 0))
    Dim SummaryPicNameList(1 To 32) As Variant
    Dim SummaryPicLocationList(1 To 32) As Variant
    Dim i As Long
    For i = 1 To 32
        If CStr(chartPicNameLocationRange.Cells(i, 1).Value) <> "" Then
            SummaryPicNameList(i) = chartPicNameLocationRange.Cells(i, 1).Value
            SummaryPicLocationList(i) = chartPicPasteLocationRange.Cells(i, 1).Value
        End If
    Next i
    Set chartSheet = ThisWorkbook.Sheets(ChartBuilderSheetName_str)
    Dim chartPic As Picture
    For i = 1 To 32
        If CStr(SummaryPicNameList(i)) <> "" Then
            Call TryCopyOBJ(chartSheet.ChartObjects(CStr(SummaryPicNameList(i))))
            t = Timer
            Set chartPic = TryPicPaste2Sheet(ThisWorkbook.Sheets(SummarySheetName_str))
            With chartPic
                .Left = ThisWorkbook.Sheets(SummarySheetName_str).Range(SummaryPicLocationList(i)).Left
                .Top = ThisWorkbook.Sheets(SummarySheetName_str).Range(SummaryPicLocationList(i)).Top
            End With
            Dim rubbish3 As Integer
            Debug.Print "》》" & SummaryPicNameList(i) & "》》统计图粘贴完成>==>耗时" & CStr(Timer - t)
        End If
    Next i
    AutoPastePicture2 = 0
    Exit Function
End Function

Function AutoPastePicture3() As Integer
    Dim t
    Dim classTitle As Range, classRangeStart As Range, classRange As Range, chartSheet As Worksheet
    Set classTitle = GetConfigTitleRange(KeyFiledsTitleLocation_str)
    Set classRange = RangeDown2Blank(classTitle.Offset(1, 0))
    Dim classNameList(1 To 32) As Variant
    Dim i As Long, noTableFlag As Boolean, currentCell As Range
    Set currentCell = Range(SummarySheetKeyFieldTitleLocation_str)
    For i = 1 To 32
        If CStr(classRange(i, 1).Value) <> "" Then
            classNameList(i) = classRange(i, 1).Value
        End If
    Next i
    Set chartSheet = ThisWorkbook.Sheets(ChartBuilderSheetName_str)
    Dim chartPic As Picture
    For i = 1 To 32
        If CStr(classNameList(i)) <> "" Then
            t = Timer
            noTableFlag = True
            Set currentCell = FindDownByRowWithStartText(currentCell, CStr(classNameList(i))).Offset(1, 0)
            Dim pic As Picture
            For Each pic In chartSheet.Pictures
                If pic.Name = CStr(classNameList(i) & "统计表") Then
                    Call TryCopyOBJ(pic)
                    noTableFlag = False
                    Exit For
                End If
            Next pic
            If noTableFlag = False Then
                Set chartPic = TryPicPaste2Sheet(ThisWorkbook.Sheets(SummarySheetName_str))
                With chartPic
                    .Left = ThisWorkbook.Sheets(SummarySheetName_str).Range(currentCell.Address).Left
                    .Top = ThisWorkbook.Sheets(SummarySheetName_str).Range(currentCell.Address).Top
                End With
                Dim rubbish3 As Integer
                rubbish3 = ClearWindowsClipboard()
                Debug.Print "》》" & CStr(classNameList(i)) & "》》表格图粘贴完成>==>耗时" & CStr(Timer - t)
            End If
            t = Timer
            Dim errorNAFlag As Boolean
            errorNAFlag = False
            On Error GoTo chartNA
                Call TryCopyOBJ(chartSheet.ChartObjects(CStr(classNameList(i) & "柱状图")))
                On Error GoTo 0
            If errorNAFlag = False Then
                Dim ramSer As Series
                Set ramSer = chartSheet.ChartObjects(CStr(classNameList(i) & "柱状图")).Chart.SeriesCollection(1)
                If ramSer.Values(1) > 0 Then
                    Set chartPic = TryPicPaste2Sheet(ThisWorkbook.Sheets(SummarySheetName_str))
                    If noTableFlag = False Then
                        With chartPic
                            .Left = ThisWorkbook.Sheets(SummarySheetName_str).Range(currentCell.Offset(0, 10).Address).Left
                            .Top = ThisWorkbook.Sheets(SummarySheetName_str).Range(currentCell.Offset(0, 10).Address).Top
                        End With
                    Else
                        With chartPic
                            .Left = ThisWorkbook.Sheets(SummarySheetName_str).Range(currentCell.Address).Left
                            .Top = ThisWorkbook.Sheets(SummarySheetName_str).Range(currentCell.Address).Top
                        End With
                    End If
                    rubbish3 = ClearWindowsClipboard()
                    Debug.Print "》》" & CStr(classNameList(i)) & "》》统计图粘贴完成>==>耗时" & CStr(Timer - t)
                Else
                    Debug.Print "》》" & CStr(classNameList(i)) & "》》统计图无有效值，不再粘贴>==>耗时" & CStr(Timer - t)
                End If
            End If
            errorNAFlag = False
        End If
    Next i
    AutoPastePicture3 = 0
    Exit Function
chartNA:
    errorNAFlag = True
    On Error GoTo 0
End Function

Function 分离饼图数据标签格式化专用函数(chartName As String) As Integer
    Dim picWorksheet As Worksheet
    Dim chtOBJ As ChartObject, targetChtOBJ As ChartObject
    Set picWorksheet = ThisWorkbook.Sheets(ChartBuilderSheetName_str)
    For Each chtOBJ In picWorksheet.ChartObjects
        If chtOBJ.Name = chartName Then
            Set targetChtOBJ = chtOBJ
            targetChtOBJ.Chart.SetElement (msoElementDataLabelBestFit)
            Exit For
        End If
    Next chtOBJ
    If targetChtOBJ Is Nothing Then
        Debug.Print "！【" & ChartBuilderSheetName_str & "】工作表中不存在名为【" & chartName & "】统计图！"
        Exit Function
    End If
    Dim points0 As Points, point_i As Point, pointCount As Long, deputyCount As Long
    Set points0 = targetChtOBJ.Chart.FullSeriesCollection(1).Points
    pointCount = points0.Count
    deputyCount = pointCount - 11
    If deputyCount > 1 Then
        targetChtOBJ.Chart.ChartGroups(1).SplitValue = deputyCount
    Else
        targetChtOBJ.Chart.ChartGroups(1).SplitValue = 1
    End If
    Dim topLocation As Long, bottomLocation As Long, leftLocation As Long, dataLabelW As Long, dataLabelH As Long
    topLocation = 40
    leftLocation = 525
    dataLabelW = 100
    dataLabelH = 15
    bottomLocation = 435 - dataLabelH - 5
    Dim cycleCount As Long
    Dim currentTopLocation As Long
    For cycleCount = 9 To pointCount
        currentTopLocation = (cycleCount - 9) * (dataLabelH + 5)
        If currentTopLocation < bottomLocation Then
            With targetChtOBJ.Chart.FullSeriesCollection(1).Points(cycleCount).DataLabel
                .ShowValue = True
                .ShowCategoryName = True
                .ShowPercentage = False
                .Width = dataLabelW
                .Height = dataLabelH
                .Left = leftLocation
                .Top = currentTopLocation
            End With
        Else
            With targetChtOBJ.Chart.FullSeriesCollection(1).Points(cycleCount).DataLabel
                .ShowValue = False
                .ShowCategoryName = False
                .ShowPercentage = False
            End With
        End If
    Next cycleCount
    分离饼图数据标签格式化专用函数 = 0
End Function

Function GetFolderPath(filePath As String) As String
    Dim folderPath As String
    folderPath = Left(filePath, InStrRev(filePath, "\") - 1)
    GetFolderPath = folderPath
End Function

Function PasteFromAnotherWorkbook(fromWorkbookName As String, fromWorksheetName As String) As Integer
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    
    Set sourceWorkbook = Workbooks.Open(fromWorkbookName)
    Set sourceSheet = sourceWorkbook.Sheets(fromWorksheetName)
    
    Set destinationWorkbook = ThisWorkbook
    Set destinationSheet = destinationWorkbook.Sheets(DetailSheetName_str)
    
    sourceSheet.UsedRange.Copy destinationSheet.Range("A1")
    
    sourceWorkbook.Close SaveChanges:=False
End Function

Function ClearWindowsClipboard() As Integer
    OpenClipboard 0
    EmptyClipboard
    CloseClipboard
    ClearWindowsClipboard = 0
End Function

Sub TryCopyOBJ(obj As Object)
    Dim try_count As Integer, sleep_mSecond As String
    Dim startTime As Double
    Dim is_copied As Boolean
    
    try_count = 0
    is_copied = False
    
    Do While try_count < 6 And Not is_copied
        If try_count <> 0 Then
            sleep_mSecond = "00:00:0" & CStr(try_count)
            Application.Wait Now + TimeValue(sleep_mSecond)
        End If
        On Error Resume Next
        obj.Copy
        If Err.Number = 0 Then
            is_copied = True
        Else
            try_count = try_count + 1
        End If
        On Error GoTo 0
    Loop
    
    If Not is_copied Then
        Err.Raise vbObjectError + 9999, "复制失败！", "！》无法完成复制操作，程序中断。"
        Exit Sub
    End If
End Sub

Function TryPicPaste2Sheet(targetSheet As Worksheet) As Picture ' 粘贴信息源自前文copy
    Dim try_count As Integer, sleep_mSecond As String
    Dim startTime As Double
    Dim is_pasted As Boolean
    
    try_count = 0
    is_pasted = False
    
    Do While try_count < 6 And Not is_pasted
        startTime = Timer
        If try_count <> 0 Then
            sleep_mSecond = "00:00:0" & CStr(try_count)
            Application.Wait Now + TimeValue(sleep_mSecond)
        End If
        On Error Resume Next
        Set TryPicPaste2Sheet = targetSheet.Pictures.Paste
        If Err.Number = 0 Then
            is_pasted = True
        Else
            try_count = try_count + 1
        End If
        On Error GoTo 0
    Loop
    
    If Not is_pasted Then
        Err.Raise vbObjectError + 9999, "粘贴失败！", "！》无法完成粘贴操作，程序中断。"
        Exit Function
    End If
End Function

Sub 一键拆分明细()
    Dim t
    t = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Dim rubbish0 As Integer
    rubbish0 = Unique2RAMSheet()
    rubbish0 = WithoutStarString()
    Dim wb As Workbook
    Dim ws0 As Worksheet
    Dim ws As Worksheet
    Dim saveSheetNameRange As Range
    Dim saveSheetName(1 To 32) As String
    Set saveSheetNameRange = RangeDown2BlankFast(GetConfigTitleRange(KeyFiledsTitleLocationShort_str).Offset(1, 0))
    Dim keywordField As String
    keywordField = GetConfigKeywordString(ConfigTitleKeyword_str)
    Dim classFullNameRange As Range
    Set classFullNameRange = RangeDown2BlankFast(GetConfigTitleRange(KeyFiledsTitleLocation_str).Offset(1, 0))
    Dim classFullName(1 To 32) As String
    Dim i As Long, b As Boolean, rubbish As Integer
    b = False
    For i = 1 To 32
        If CStr(saveSheetNameRange(i, 1).Value) <> "" Then
            saveSheetName(i) = CStr(saveSheetNameRange(i, 1).Value)
            classFullName(i) = CStr(classFullNameRange(i, 1).Value)
        End If
    Next i
    Set wb = ThisWorkbook
    Set ws0 = wb.Sheets(DetailSheetName_str)
    For i = 1 To 32
        If CStr(saveSheetNameRange(i, 1).Value) <> "" Then
            Dim t1
            t1 = Timer
            Set ws = Nothing
            For Each ws In wb.Sheets
                If ws.Name = CStr(saveSheetNameRange(i, 1).Value) Then
                    b = True
                End If
            Next ws
            If b = True Then
                Application.DisplayAlerts = False
                wb.Sheets(CStr(saveSheetNameRange(i, 1).Value)).Delete
                Application.DisplayAlerts = True
            End If
            Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count - 7))
            ws.Name = saveSheetName(i)
            Dim lastRow As Long
            Dim columnIndex As Long
            columnIndex = Application.Match(keywordField, ws0.Rows(1), 0)
            Dim i1 As Long
            Dim j As Long
            lastRow = ws0.Cells(ws0.Rows.Count, columnIndex).End(xlUp).Row
            ws0.Rows(1).Copy ws.Rows(1) ' 复制标题行
            For i1 = 2 To lastRow
                If ws0.Cells(i1, columnIndex).Value = classFullName(i) Then
                    j = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row + 1
                    ws0.Rows(i1).Copy ws.Rows(j)
                End If
            Next i1
            Debug.Print "》》" & saveSheetName(i) & "》》脱离拆分完成>==>耗时" & CStr(Timer - t1)
        End If
        b = False
    Next i
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ThisWorkbook.Worksheets(ChartBuilderSheetName_str).Select
    Debug.Print "》" & DetailSheetName_str & "》细分拆分全部完成>==>耗时" & CStr(Timer - t)
End Sub

Sub 制作图表()
    Calculate
    Dim t
    t = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Dim rubbish0 As Integer
    ThisWorkbook.Worksheets(ChartBuilderSheetName_str).Range("A1").Select
    rubbish0 = ResetDataSourceAuto("总图")
    rubbish0 = 分离饼图数据标签格式化专用函数("总图")
    rubbish0 = ResetDataSourceAuto("考核土地饼状图")
    rubbish0 = ResetDataSourceAuto("考核土地柱状图")
    rubbish0 = 城市分布生成图专用函数("城市分布柱状图")
    Dim chartSheet As Worksheet
    Set chartSheet = ThisWorkbook.Sheets(ChartBuilderSheetName_str)
    Dim classTitle As Range, classRangeStart As Range, classRange As Range
    Set classTitle = GetConfigTitleRange(ChartBuliderSourceLocation_str)
    Set classRangeStart = classTitle.Offset(1, 0)
    Dim classRangeStartSheetStr As String, classRangeStartCellStr As String
    Dim classRangeStrS() As String
    classRangeStrS = Split(classRangeStart, "!")
    classRangeStartSheetStr = classRangeStrS(0)
    classRangeStartCellStr = classRangeStrS(1)
    Set classRange = RangeDown2Blank(classRangeStart)
    Dim classificationAddress As Range
    Dim forEachCount As Long
    forEachCount = 0
    For Each classificationAddress In classRange
        Dim t1
        t1 = Timer
        Dim sourceRange As Range
        Dim classificationAddressStr As String
        classificationAddressStr = CStr(classificationAddress.Value)
        Set sourceRange = Range(classificationAddressStr)
        Set sourceRange = RangeDown2Blank(sourceRange)
        Set sourceRange = sourceRange.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count + 2)
        Dim classificationChart As ChartObject
        Dim topOffset As Long
        topOffset = forEachCount * 310 + 750
        If sourceRange.Rows.Count > 20 Then
            Set classificationChart = chartSheet.ChartObjects.Add(Left:=0, Top:=topOffset, Width:=1310, Height:=300)
            '下面的属性是只读的，只能select了
            classificationChart.Chart.PlotArea.Select
            Selection.Width = 1258
            Selection.Left = 65
        Else
            Set classificationChart = chartSheet.ChartObjects.Add(Left:=600, Top:=topOffset, Width:=710, Height:=300)
        End If
        Dim cht As ChartObject
        For Each cht In chartSheet.ChartObjects
            If cht.Name = CStr(classificationAddress.Offset(0, -1).Value) & "柱状图" Then
                cht.Delete
                Exit For
            End If
        Next cht
        Dim shp As Shape
        For Each shp In chartSheet.Shapes
            If shp.Name = CStr(classificationAddress.Offset(0, -1).Value) & "统计表" Then
                shp.Delete
                Exit For
            End If
        Next shp
        classificationChart.Name = CStr(classificationAddress.Offset(0, -1).Value) & "柱状图"
        With classificationChart.Chart
            .SetSourceData Source:=sourceRange
            .PlotBy = xlColumns
            .SetElement (msoElementDataLabelOutSideEnd)
            .ChartType = xlColumnClustered
            .FullSeriesCollection(1).ChartType = xlColumnClustered
            .FullSeriesCollection(1).AxisGroup = 1
            .FullSeriesCollection(2).ChartType = xlLine
            .FullSeriesCollection(2).AxisGroup = 1
            .FullSeriesCollection(2).AxisGroup = 2
            .SetElement (msoElementChartTitleAboveChart)
            .SetElement (msoElementLegendBottom)
            .FullSeriesCollection(2).ApplyDataLabels
            .FullSeriesCollection(2).DataLabels.Position = xlLabelPositionAbove
        End With
        If sourceRange.Rows.Count > 55 Then
            classificationChart.Chart.FullSeriesCollection(2).DataLabels.Delete
        End If
        If sourceRange.Rows.Count > 100 Then
            classificationChart.Chart.FullSeriesCollection(1).DataLabels.Delete
        End If
        classificationChart.Chart.ChartTitle.text = (CStr(classificationAddress.Offset(0, -1).Value) & "分布")
        If sourceRange.Rows.Count <= 20 Then
            Dim tablePicSourceRange As Range
            Dim tablePic As Picture
            Set tablePicSourceRange = sourceRange
            Dim ws2 As Worksheet
            Set ws2 = sourceRange.Parent
            ws2.Activate
            Call TryCopyOBJ(tablePicSourceRange)
            Set tablePic = TryPicPaste2Sheet(ThisWorkbook.Sheets(ChartBuilderSheetName_str))
            With tablePic
                .Left = 0
                .Top = topOffset
                .Height = 300
                .Name = (CStr(classificationAddress.Offset(0, -1).Value) & "统计表")
            End With
            If tablePic.Width > 500 Then
                With tablePic
                    .Width = 500
                End With
            End If
        End If
        Dim rubbish3 As Integer
        rubbish3 = ClearWindowsClipboard()
        forEachCount = forEachCount + 1
        Debug.Print "》》" & CStr(classificationAddress.Offset(0, -1).Value) & "》》图表生成完成>==>耗时" & CStr(Timer - t1)
    Next classificationAddress
    ThisWorkbook.Worksheets(ChartBuilderSheetName_str).Select
    Debug.Print "》图表生成全部完成>==>耗时" & CStr(Timer - t)
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub 寻找并格式化概述文本()
    Dim startCell As Range, targetCell As Range, forCount As Integer
    Set startCell = Range(SummarySheetKeyFieldTitleLocation_str)
    Set targetCell = startCell
    For forCount = 1 To (RangeDown2Blank(GetConfigTitleRange(ChartBuliderSourceLocation_str)).Rows.Count - 1)
        Set targetCell = FindDownByRowWithStartText(targetCell, OverviewKeywords_str)
        Dim rubbish As Integer
        rubbish = CopyAndPasteAsValue(targetCell)
        rubbish = FormatText_GS(targetCell)
        On Error GoTo targetCellNA
            Set targetCell = targetCell.Offset(1, 0)
    Next forCount
    ThisWorkbook.Worksheets(ChartBuilderSheetName_str).Range("A1").Select
targetCellNA:
    Debug.Print ("！寻找并格式化概述文本()的targetCell出现Nothing，已跳过")
End Sub
Sub 一键粘贴图片()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Dim t
    t = Timer
    Dim rubbish As Integer
    Dim ws As Worksheet, currentWs As Worksheet
    Dim pic As Shape
    Set currentWs = ActiveSheet
    Set ws = ThisWorkbook.Worksheets(SummarySheetName_str)
    For Each pic In ws.Shapes
        pic.Delete
    Next pic
    ws.Activate
    ws.Parent.Windows(1).Zoom = 100
    currentWs.Activate
    rubbish = AutoPastePicture1()
    rubbish = AutoPastePicture2()
    rubbish = AutoPastePicture3()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    ThisWorkbook.Worksheets(ChartBuilderSheetName_str).Select
    Debug.Print "》图表粘贴全部完成>==>耗时" & CStr(Timer - t)
End Sub

Sub 一键兼容性保存()
    Dim path As String
    path = ThisWorkbook.path
    Call 递归一键兼容性保存(path)
End Sub

Sub 一键操作全家桶()
    Dim t
    t = Timer
    Debug.Print "=+=+=+=+=+=+=+=+=+=+=  一键操作开始执行  =+=+=+=+=+=+=+=+=+=+="
    Call 一键拆分明细
    Call 制作图表
    Call 寻找并格式化概述文本
    Call 一键粘贴图片
    Call 一键兼容性保存
    Debug.Print Chr(10) & "所有任务全部完成>==>耗时" & CStr(Timer - t)
    Debug.Print "SUCCESSFUL_COMPLETION"
End Sub

Sub 递归一键兼容性保存(targetSavePath As String)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Dim t
    t = Timer
    Dim path As String
    path = targetSavePath
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim ws As Worksheet
    Dim saveSheetNameRange As Range
    Dim saveSheetName(1 To 32) As String
    Set saveSheetNameRange = RangeDown2Blank(GetConfigTitleRange(KeyFiledsTitleLocationShort_str))
    Dim i As Long
    For i = 1 To 32
        If CStr(saveSheetNameRange(i, 1).Value) <> "" Then
            saveSheetName(i) = CStr(saveSheetNameRange(i, 1).Value)
        End If
    Next i
    Set wb = ThisWorkbook
    Set newWb = Workbooks.Add
    For Each ws In wb.Sheets
        If ws.Name = SummarySheetName_str Then
            ws.Copy After:=newWb.Sheets(newWb.Sheets.Count)
            Exit For
        End If
    Next ws
    Sheets("Sheet1").Select
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    For Each ws In wb.Sheets
        If ws.Name = DetailSheetName_str Then
            ws.Copy After:=newWb.Sheets(newWb.Sheets.Count)
            Exit For
        End If
    Next ws
    For i = 1 To 32
        If CStr(saveSheetNameRange(i, 1).Value) <> "" Then
            For Each ws In wb.Sheets
                If ws.Name = CStr(saveSheetNameRange(i, 1).Value) Then
                    ws.Copy After:=newWb.Sheets(newWb.Sheets.Count)
                End If
            Next ws
        End If
    Next i
    Sheets(SummarySheetName_str).Select
    
    Dim links As Variant
    links = newWb.LinkSources(Type:=xlLinkTypeExcelLinks)
    
    If Not IsEmpty(links) Then
        Dim j As Long
        For j = 1 To UBound(links)
            ActiveWorkbook.BreakLink Name:=links(j), Type:=xlLinkTypeExcelLinks
        Next j
    End If
    newWb.SaveAs path & "\00" & Range(WORKBOOK_NAME_LOCATION_str).Value & ".xlsx"
    newWb.Close

    Set newWb = Nothing
    Set wb = Nothing
    Debug.Print "》" & path & "\00" & Range(WORKBOOK_NAME_LOCATION_str).Value & ".xlsx》兼容性输出完成>==>耗时" & CStr(Timer - t)
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub 递归一键操作全家桶()
    
    Dim t, t1
    t = Timer
    Dim rubbish As Integer
    Dim fso As Object, file As Object
    Dim filePath As String, textLine As String, targetSavePath As String
    Dim ws As Worksheet
    
    ' 创建FileSystemObject对象
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = ThisWorkbook.path & "\path_details.txt"
    Set file = fso.OpenTextFile(filePath, 1) ' 1表示只读模式
    
    Do While Not file.AtEndOfStream
        t1 = Timer
        textLine = file.ReadLine
        Debug.Print "》当前正在处理的文件是：" & textLine
        Set ws = ThisWorkbook.Sheets(DetailSheetName_str)
        ws.Cells.ClearContents
        rubbish = PasteFromAnotherWorkbook(textLine, "Sheet1")
        
        Call 一键拆分明细
        Call 制作图表
        Call 寻找并格式化概述文本
        Call 一键粘贴图片
        targetSavePath = GetFolderPath(textLine)
        Call 递归一键兼容性保存(targetSavePath)
        Application.Wait Now + TimeValue("00:00:02")
        
        Debug.Print Chr(10) & "当前遍历对象任务完成>==>耗时" & CStr(Timer - t1)
    Loop
    
    file.Close
    Set file = Nothing
    Set fso = Nothing
    Debug.Print Chr(10) & "递归任务全部完成>==>耗时" & CStr(Timer - t)
End Sub

Sub 远程一键递归()
    Application.Wait Now + TimeValue("00:00:20")
    Call 递归一键操作全家桶
End Sub
