Attribute VB_Name = "Module1"
Option Explicit

Sub CopyFromCSV()
Attribute CopyFromCSV.VB_ProcData.VB_Invoke_Func = "C\n14"

    Dim xFd As FileDialog
    Dim xSPath As String
    Dim xCSVFile As String
    Dim xWsheet As String
    Dim wbCopy As Workbook
    Dim wsCopy As Worksheet
    Dim wsDest As Worksheet
    Dim rowsDest As Integer
    Dim result As String
    
    Application.DisplayAlerts = False
    Application.StatusBar = True
    
    xWsheet = ActiveWorkbook.Name
    Workbooks(xWsheet).Sheets.Add.Name = "Results"
    
    Set xFd = Application.FileDialog(msoFileDialogFolderPicker)
    xFd.title = "Select a folder:"
    
    If xFd.Show = -1 Then
        xSPath = xFd.SelectedItems(1)
    Else
        Exit Sub
    End If
    
    If Right(xSPath, 1) <> "\" Then xSPath = xSPath + "\"
    
    xCSVFile = Dir(xSPath & "*.csv")
    
    Do While xCSVFile <> ""
        Application.StatusBar = "Copying: " & xCSVFile
        
        Set wsDest = Workbooks(xWsheet).Sheets("Results")
        Set wbCopy = Workbooks.Open(xSPath & xCSVFile)
        Set wsCopy = wbCopy.Worksheets(1)
        
        rowsDest = wsDest.Range("A" & Rows.Count).End(xlUp).row
        
        wsDest.Range("A" & rowsDest + 2).Value2 = Replace(xCSVFile, ".csv", "")
        wsCopy.UsedRange.Copy wsDest.Range("A" & rowsDest + 3)
        
        wbCopy.Close
        
        xCSVFile = Dir
        
    Loop
    
    result = Workbooks(xWsheet).Sheets("Results").UsedRange.Replace(What:=",", Replacement:=".", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True)
    result = Workbooks(xWsheet).Sheets("Results").Range("A:A").Replace(What:="high", Replacement:="full", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True)
    
    Application.StatusBar = False
    Application.DisplayAlerts = True
End Sub

Sub SortOutData()
Attribute SortOutData.VB_ProcData.VB_Invoke_Func = "S\n14"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ws_sycl As Worksheet
    Dim ws_opencl As Worksheet
    Dim ws_gcc As Worksheet
    Dim last_row As Integer
    Dim last_column As Integer
    Dim counter As Integer
    Dim region_begin As Integer
    Dim region_end As Integer
    Dim sycl, opencl, gcc As Boolean
    
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets("Results")
    
    last_row = ws.UsedRange.SpecialCells(xlCellTypeLastCell).row
    last_column = ws.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    wb.Sheets.Add.Name = "SYCL"
    wb.Sheets.Add.Name = "OpenCL"
    wb.Sheets.Add.Name = "GCC"
    
    Set ws_sycl = wb.Worksheets("SYCL")
    Set ws_opencl = wb.Worksheets("OpenCL")
    Set ws_gcc = wb.Worksheets("GCC")
    
    counter = 1
    
    Do Until counter >= last_row
    
    region_begin = ws.Range("A" & counter).End(xlDown).row
    
    sycl = InStr(ws.Range("A" & region_begin).Value2, "sycl") > 0
    opencl = InStr(ws.Range("A" & region_begin).Value2, "opencl") > 0
    gcc = InStr(ws.Range("A" & region_begin).Value2, "gcc") > 0
    
    region_end = ws.Range("A" & region_begin).End(xlDown).row
    
    If sycl = True Then
        Copy_Data ws_sycl, ws, counter, region_begin, region_end, last_column
    End If
    
    If opencl = True Then
        Copy_Data ws_opencl, ws, counter, region_begin, region_end, last_column
    End If
    
    If gcc = True Then
        Copy_Data ws_gcc, ws, counter, region_begin, region_end, last_column
    End If
    
    counter = region_end + 1
    
    Loop
    
End Sub

Sub AnalyseData()
Attribute AnalyseData.VB_ProcData.VB_Invoke_Func = "A\n14"
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim ws_result As Worksheet
    Dim ws_sycl As Worksheet
    Dim ws_opencl As Worksheet
    Dim ws_gcc As Worksheet
    Dim sycl_begin As Integer
    Dim sycl_end As Integer
    Dim opencl_begin As Integer
    Dim opencl_end As Integer
    Dim gcc_begin As Integer
    Dim gcc_end As Integer
    Dim counter As Integer
    Dim project_begin As Integer
    Dim project_end As Integer
    Dim i, k, last_col As Integer
    Dim header As Variant
    Dim table_header As Variant
    
    Set wb = ActiveWorkbook
    wb.Sheets.Add.Name = "Analysis"
    
    Set ws = wb.Worksheets("Analysis")
    Set ws_result = wb.Worksheets("Results")
    Set ws_sycl = wb.Worksheets("SYCL")
    Set ws_opencl = wb.Worksheets("OpenCL")
    Set ws_gcc = wb.Worksheets("GCC")
    
    'Copy Project Names
    project_begin = ws_result.Range("A:A").Find("Project", MatchCase:=False).row
    project_end = ws_result.Range("A" & project_begin).End(xlDown).row
    
    counter = ws.Range("A" & Rows.Count).End(xlUp).row
    
    header = Array("GMX Performance (ns/day)", "GMX Wall time (s)", "CPU_Usage", "GPU_Usage", "CPU_Freq", "GPU_Freq")
    table_header = Array("GMX Performance (ns/day)", "GMX Wall time (s)", "CPU_Usage", "GPU_Usage", "CPU_Core_All_Avg_Freq", "GPU_Freq_act")
    
    EnterData ws, ws_result, counter, last_col, project_begin, project_end, header, table_header, "SYCL", "sycl"
    
    'OpenCL
    
    counter = ws.Range("A" & Rows.Count).End(xlUp).row + 25
    
    EnterData ws, ws_result, counter, last_col, project_begin, project_end, header, table_header, "OpenCL", "opencl"
    
    'GCC
    
    counter = ws.Range("A" & Rows.Count).End(xlUp).row + 25
    
    EnterData ws, ws_result, counter, last_col, project_begin, project_end, header, table_header, "GCC", "gcc"

End Sub

Sub EnterData(ws As Worksheet, ws_result As Worksheet, counter As Integer, last_col As Integer, project_begin As Integer, _
                project_end As Integer, header As Variant, table_header As Variant, title As String, table_name As String)
    
    Dim i, k As Integer
    Dim result As String
    
    ws.Cells(counter, 1).Value2 = "GMX_" & title

    For i = 1 To 6
        ws.Cells(counter + 1, (3 * i) - 1).Value2 = header(i - 1)
        ws.Range(Cells(counter + 1, (3 * i) - 1), Cells(counter + 1, (3 * i) + 1)).Merge
        ws.Range(Cells(counter + 1, (3 * i) - 1), Cells(counter + 1, (3 * i) + 1)).HorizontalAlignment = xlCenter
    Next i
    
    last_col = ws.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    For i = 2 To 19 Step 3
        ws.Cells(counter + 2, i).Value2 = "Full"
        ws.Cells(counter + 2, i + 1).Value2 = "Medium"
        ws.Cells(counter + 2, i + 2).Value2 = "Low"
    Next i
    
    last_col = ws.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    
    ws_result.Activate
    ws_result.Range(Cells(project_begin, 1), Cells(project_end, 1)).Copy ws.Range("A" & counter + 2)
    ws.Activate
    
    k = 0
    
    For i = 2 To last_col Step 3
        ws.Cells(counter + 3, i).Value2 = "=" & table_name & "_full[" & table_header(k) & "]"
        ws.Cells(counter + 3, i + 1).Value2 = "=" & table_name & "_medium[" & table_header(k) & "]"
        ws.Cells(counter + 3, i + 2).Value2 = "=" & table_name & "_low[" & table_header(k) & "]"
        k = k + 1
    Next i
    
    result = ws.UsedRange.Replace(What:="@", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False, _
        FormulaVersion:=xlReplaceFormula2)
        
    'Create_Performance_Chart counter + 1, counter + (project_end - project_begin) + 2, "GMX_" & title & " Performance (ns/day)"
    
        
End Sub

Sub Copy_Data(ws As Worksheet, ws_result As Worksheet, counter As Integer, region_begin As Integer, _
                region_end As Integer, last_column As Integer)
    
    Dim table_name As String
    Dim name_array As Variant
    Dim row_nr As Integer
    
    row_nr = ws.Range("A" & counter).End(xlUp).row
    ws_result.Activate
    ws_result.Range(Cells(region_begin, 1), Cells(region_end, last_column)).Copy ws.Range("A" & row_nr + 2)
    ws.Activate
    name_array = Split(ws.Range("A" & row_nr + 2), "_")
    table_name = name_array(1) & "_" & name_array(2)
    ws.ListObjects.Add(xlSrcRange, ws.Range(Cells(row_nr + 3, 1), _
    Cells(row_nr + 3 + (region_end - region_begin - 1), last_column)), , xlYes).Name = table_name
    
End Sub

Sub Create_Performance_Chart(range_start As Integer, range_end As Integer, title As String)
    
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets("Analysis")
    
    ws.Range("A" & range_start & ":D" & range_end).Select
    ws.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=ws.Range("A" & range_start & ":D" & range_end)
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = title
    ActiveChart.ChartGroups(1).FullCategoryCollection(1).IsFiltered = True

End Sub

Sub Create_CPU_GPU_Utilisation_Chart(range_start As Integer, range_end As Integer, title As String)
    
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.Worksheets("Analysis")
    
    ws.Range("A" & range_start & ":M" & range_end).Select
    ws.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=ws.Range("A" & range_start & ":M" & range_end)
    'ActiveChart.ChartGroups(1).FullCategoryCollection(1).IsFiltered = True
    ActiveChart.FullSeriesCollection(1).IsFiltered = True
    ActiveChart.FullSeriesCollection(2).IsFiltered = True
    ActiveChart.FullSeriesCollection(3).IsFiltered = True
    ActiveChart.FullSeriesCollection(4).IsFiltered = True
    ActiveChart.FullSeriesCollection(5).IsFiltered = True
    ActiveChart.FullSeriesCollection(6).IsFiltered = True
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(7).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(7).AxisGroup = 1
    ActiveChart.FullSeriesCollection(8).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(8).AxisGroup = 1
    ActiveChart.FullSeriesCollection(9).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(9).AxisGroup = 1
    ActiveChart.FullSeriesCollection(10).ChartType = xlLine
    ActiveChart.FullSeriesCollection(10).AxisGroup = 1
    ActiveChart.FullSeriesCollection(11).ChartType = xlLine
    ActiveChart.FullSeriesCollection(11).AxisGroup = 1
    ActiveChart.FullSeriesCollection(12).ChartType = xlLine
    ActiveChart.FullSeriesCollection(12).AxisGroup = 1
    ActiveChart.ChartColor = 13
    ActiveChart.ChartTitle.Text = title
    
End Sub

