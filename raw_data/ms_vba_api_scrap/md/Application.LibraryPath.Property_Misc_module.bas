Attribute VB_Name = "Misc_module"
Public Sub GetValueFromBrowser()
    Dim ie As Object
    Dim url As String
    Dim myPoints As String

    url = "https://mazemi.github.io/Add-ins_guide/"
    Set ie = CreateObject("InternetExplorer.Application")

    With ie
      .Visible = 0
      .Navigate url
       While .Busy Or .ReadyState <> 4
         DoEvents
       Wend
    End With
Debug.Print 1
    Dim Doc As HTMLDocument
    Set Doc = ie.Document

    myPoints = Trim(Doc.getElementsByName("ram-version")(0).Value)
    Range("A1").Value = myPoints
Debug.Print myPoints
End Sub

Sub show_sheet()
    sheets("dissagregation_setting").Visible = True
    sheets("analysis_list").Visible = True
End Sub

Sub compare_strata()
    Dim r1 As Range
    Dim r2 As Range
    Dim res As Boolean
    Dim d As Variant
    
    a = Cells(rows.count, 1).End(xlUp).Row
    b = Cells(rows.count, 2).End(xlUp).Row

    Dim col1 As New Collection
    Dim col2 As New Collection
    
    Set r1 = sheets("inpro").Range("A2:A" & a)
    Set r2 = sheets("inpro").Range("B2:B" & b)
    
    For Each i In r1
        col1.Add CStr(i)
    Next
    
    For Each j In r2
        col2.Add CStr(j)
    Next
    
    For Each d In col2
        res = HasKey(col1, CStr(d))
        If Not res Then
'            Debug.Print d
            
        End If
    Next
    Debug.Print "done!"
End Sub

Function HasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    HasKey = (err.Number = 0)
    err.Clear
End Function


Sub show_last()
    On Error Resume Next
    Dim header_arr() As Variant
    
    last_row = sheets(find_main_data).Cells(rows.count, 1).End(xlUp).Row
    last_row2 = sheets(find_main_data).UsedRange.rows(ActiveSheet.UsedRange.rows.count).Row
    last_col = sheets(find_main_data).Cells(1, columns.count).End(xlToLeft).Column
    
    ' below needs to be improved
    header_arr = sheets(find_main_data).Range(Cells(1, 1), Cells(1, 1).End(xlToRight)).Value2
    Debug.Print last_row, last_row2, last_col, LBound(header_arr), UBound(header_arr)
    
End Sub

Public Function CellType(c)
    On Error Resume Next
    Application.Volatile
    Select Case True
    Case IsEmpty(c): CellType = "Blank"
    Case Application.IsText(c): CellType = "Text"
    Case IsNumeric(CInt(c)): CellType = "Number"
        
    Case Application.IsLogical(c): CellType = "Logical"
    Case Application.IsErr(c): CellType = "Error"
    Case IsDate(c): CellType = "Date"
    End Select
End Function

Sub ShowLibraryPaths()
    MsgBox "Library Path: " & Application.LibraryPath & vbCrLf & _
           "User Library Path: " & Application.UserLibraryPath, vbOKOnly
End Sub

Sub Convert_to_Text(ByRef xRange As String, Optional ByVal W_Sheet As Worksheet)
    Dim TP As Double
    Dim V_Range As Range
    Dim xCell As Object
    If W_Sheet Is Nothing Then Set W_Sheet = ActiveSheet
    Set V_Range = W_Sheet.Range(xRange).SpecialCells(xlCellTypeVisible)
    For Each xCell In V_Range
        If Not IsEmpty(xCell.Value) And IsNumeric(xCell.Value) Then
            TP = xCell.Value
            xCell.ClearContents
            xCell.NumberFormat = "@"
            xCell.Value = CStr(TP)
        End If
    Next xCell
End Sub

Sub show_keen()
    If worksheet_exists("keen") Then
        sheets("keen").Visible = True
    End If
End Sub

Function MeanOfColumn(arr As Variant, colIndex As Long) As Double
    Dim i As Long
    Dim colArr() As Double
    ReDim colArr(1 To UBound(arr, 1))
    For i = 1 To UBound(arr, 1)
        colArr(i) = arr(i, colIndex)
    Next i
    MeanOfColumn = WorksheetFunction.Average(colArr)
End Function

Function MedianOfColumn(arr As Variant, colIndex As Long) As Double
    Dim i As Long
    Dim colArr() As Double
    ReDim colArr(1 To UBound(arr, 1))
    For i = 1 To UBound(arr, 1)
        colArr(i) = arr(i, colIndex)
    Next i
    MedianOfColumn = WorksheetFunction.median(colArr)
End Function


