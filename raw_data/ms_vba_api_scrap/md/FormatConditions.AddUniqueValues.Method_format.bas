Attribute VB_Name = "format"
Sub formatAll()
Attribute formatAll.VB_ProcData.VB_Invoke_Func = "k\n14"
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
   
    With Cells
'       .Interior.Color = xlNone
        .Font.Size = 10                                             ' Font Size to 11
        .Font.Name = "Arial"                                      ' Font to Calibri. Can change to any font you want
'       .FormatConditions.delete                                    ' Deletes conditional formatting
'       .WrapText = False                                           ' Unwraps all text so it's all on 1 line space
'       .Validation.delete                                          ' Removes dumb Data validation and its drop-down menu, specifically found in 'Date'
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With
    
    Rows(1).Interior.Color = RGB(0, 0, 0)                     ' Highlights Row 1 black
    Rows(1).Font.Color = RGB(255, 255, 255)
    Rows(1).Font.Bold = True                                        ' Bolds Row 1
' ----------------------------------------------------------------------------------------
    Set MR = Range("A1", Cells(1, Columns.count).End(xlToLeft))
    For Each Cell In MR                                             ' Should format date as mm/dd/yyyy. Date (with time)"
        If InStr(1, Cell.Value, "Date") Then
            Cell.Value = "Registration Date"
            Cell.EntireColumn.NumberFormat = "mm/dd/yyyy"
            Exit For
        End If
    Next
' ----------------------------------------------------------------------------------------
    ActiveSheet.Columns.AutoFit                                     ' Auto-fits/auto-spaces rows and columns
    ActiveSheet.Rows.AutoFit
' ----------------------------------------------------------------------------------------
'Condition Format Unique recorded from macro recorder
    Columns("A:A").FormatConditions.AddUniqueValues
    Columns("A:A").FormatConditions(Columns("A:A").FormatConditions.count).SetFirstPriority
    Columns("A:A").FormatConditions(1).DupeUnique = xlDuplicate
    With Columns("A:A").FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Columns("A:A").FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Columns("A:A").FormatConditions(1).StopIfTrue = False
    
    Application.ScreenUpdating = True
    Cells(1, 1).Select
    
    
    Application.DisplayStatusBar = True
    
End Sub
