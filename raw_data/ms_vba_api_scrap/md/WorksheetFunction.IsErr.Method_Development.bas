Attribute VB_Name = "Development"
'Development.bas
'Subroutines/UDF Functions currently in development; writing, testing, and debugging.
'********************************************************************************************************************************************
'--------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------- Module Settings and Global Variables ----------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'********************************************************************************************************************************************
Option Explicit
Public CSV_Home As Range
'********************************************************************************************************************************************
'--------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------------ UDF Cell Modification Workarounds ---------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'********************************************************************************************************************************************
Function SetIt(refCell)
    refCell.Parent.Evaluate "SetColor(" & refCell.Address(False, False) & ")"
    refCell.Parent.Evaluate "SetValue(" & refCell.Address(False, False) & ")"
    refCell.Parent.Evaluate "AddName(" & refCell.Address(False, False) & ")"

    MsgBox Application.EnableEvents
    refCell.Parent.Evaluate "ChangeEvents(" & refCell.Address(False, False) & ")"
    MsgBox Application.EnableEvents

    SetIt = ""
End Function

'~~> Format cells on the spreadsheet.
Sub SetColor(refCell As Range)
    refCell.Interior.ColorIndex = 3 '<~~ Change color to red
End Sub

'~~> Change another cell's value.
Sub SetValue(refCell As Range)
   refCell.Offset(, 1).value = "Sid"
End Sub

'~~> Add names to a workbook.
Sub AddName(refCell As Range)
   refCell.Name = "Sid"
End Sub

'~~> Change events
Sub ChangeEvents(refCell As Range)
    Application.EnableEvents = False
Public Function currentCaller() As String
    If TypeOf Application.Caller Is Range Then
        Dim rng As Range
        Set rng = Application.Caller '''

        currentCaller = rng.Address(External:=True)
    Else
        currentCaller = CStr(Application.Caller)
    End If
End Function
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'------------------------------------------ UDF Cell Modification Workarounds  ------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'https://stackoverflow.com/questions/12501759/vba-update-other-cells-via-user-defined-function
'------------------------------------------------------------------------------------------------------------------
'Public triggger As Boolean
'Public carryover As Variant
''
'Function reallysimple(r As Range) As Variant
'    triggger = True
'    reallysimple = r.Value
'    carryover = r.Value / 99
'End Function
''In worksheet code:
'
'Private Sub Worksheet_Calculate()
'    If Not triggger Then Exit Sub
'    triggger = False
'    Range("C1").Value = carryover
'End Sub
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'------------------------------------------ UDF Cell Modification Workarounds  ------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
Public Function MyCustomFunc()
    MsgBox Application.Caller.Parent.Name & ", " & Application.Caller.Parent.Name & ", " & Application.Caller.Address
End Function
Public Function Thumb()
    '~~> Judicious use of 'Volatile' is advised.
    '~~> This will be called everytime the sheet is recalculated
    Application.Volatile
    Application.ScreenUpdating = False
    
    MsgBox Application.Caller.Parent.Name & ", " & Application.Caller.Parent.Name & ", " & Application.Caller.Address
    
    Dim Pshp As Shape
    On Error Resume Next
    Dim PixUnitsConversion As Double
    Dim ImagePixels As Integer
    ImagePixels = 50
    PixUnitsConversion = 8.43 / 64
    Set rng = ActiveSheet.Range("J5:J36")
    Range("K1").ColumnWidth = PixUnitsConversion * ImagePixels
    '13.171875
    Range("K5:K36").RowHeight = ImagePixels
    For Each cell In rng
        filenam = cell
        ActiveSheet.Pictures.Insert(filenam).Select
        Set Pshp = Selection.ShapeRange.Item(1)
        With Pshp
           .LockAspectRatio = msoTrue
           .Width = ImagePixels
           .Height = ImagePixels
           .Cut
        End With
        Cells(cell.row, cell.Column + 1).PasteSpecial
        'Range("K1").ColumnWidth = 5 * Pshp.Height
    Next
    Application.ScreenUpdating = True
    'Application.Calculation = xlCalculationAutomatic
    'Application.DisplayAlerts = True
End Function
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'------------------------------------------ UDF Cell Modification Workarounds  ------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
'Public Function SplitAllFunc(xRg As Range)
Public Function SplitAllFunc()
    Dim xRg, xRg1, xCell As Range, i As Long, xAddress As String, xUpdate As Boolean, xRet As Variant
    On Error Resume Next
    xAddress = Application.ActiveWindow.RangeSelection.Address
    'Set xRg = Application.InputBox("Please select a range", "Kutools for Excel", xAddress, , , , , 8)
    Set xRg = Application.InputBox("Please select a range", "Kutools for Excel", xAddress, , , , , 8)
    Set xRg = Application.Intersect(xRg, xRg.Worksheet.UsedRange)
    If xRg Is Nothing Then Exit Function
        If xRg.Columns.count > 1 Then
            MsgBox "You can't select multiple columns", , "Kutools for Excel"
            Exit Function
        End If
        Set xRg1 = Application.InputBox("Split to (single cell):", "Kutools for Excel", , , , , , 8)
        Set xRg1 = xRg1.Range("A1")
        If xRg1 Is Nothing Then Exit Function
        xUpdate = Application.ScreenUpdating
        Application.ScreenUpdating = False
        For Each xCell In xRg
            xRet = Split(xCell.value, ",")
            xRg1.Worksheet.Range(xRg1.Offset(i, 0), xRg1.Offset(i + UBound(xRet, 1), 0)) = Application.WorksheetFunction.Transpose(xRet)
            i = i + UBound(xRet, 1) + 1
        Next
        Application.ScreenUpdating = xUpdate
End Function
Public Function FuncDefTest(refCell)
'Public Function FuncDefTest()
    'FuncDefTest = RefCell.Address & ", " & RefCell.Address(True, True) & ", " & RefCell.Address(False, False) & ", " & RefCell.Address(True, False) & ", " & RefCell.Address(False, True) & ", " & Application.ThisCell.Address & ", " & Application.Caller.Parent.Name & ", " & Application.Caller.Parent.Name & ", " & ActiveCell.Row
    Dim xRg As Range
    Dim xRg1 As Range
    Dim xCell As Range
    Dim i As Long
    Dim xAddress As String
    Dim xUpdate As Boolean
    Dim xRet As Variant
    On Error Resume Next
    xAddress = Application.ActiveWindow.RangeSelection.Address
    Set xRg = Application.InputBox("Please select a range", "Kutools for Excel", xAddress, , , , , 8)
    'Set xRg = Application.ThisCell
    'Set xRg = RefCell
    Set xRg = Application.Intersect(xRg, xRg.Worksheet.UsedRange)
    If xRg Is Nothing Then Exit Function
        If xRg.Columns.count > 1 Then
            MsgBox "You can't select multiple columns", , "Kutools for Excel"
            Exit Function
        End If
        Set xRg1 = Application.InputBox("Split to (single cell):", "Kutools for Excel", , , , , , 8)
        'Set xRg1 = RefCell
        'Set xRg1 = Application.ThisCell
        Set xRg1 = xRg1.Range("A1")
        If xRg1 Is Nothing Then Exit Function
        xUpdate = Application.ScreenUpdating
        Application.ScreenUpdating = False
        For Each xCell In xRg
            xRet = Split(xCell.value, ",")
            xRg1.Worksheet.Range(xRg1.Offset(i, 0), xRg1.Offset(i + UBound(xRet, 1), 0)) = Application.WorksheetFunction.Transpose(xRet)
            i = i + UBound(xRet, 1) + 1
        Next
        Application.ScreenUpdating = xUpdate
End Function
'******************************************************************************************************************
'------------------------------------------------------------------------------------------------------------------
'------------------------------------------ UDF Cell Modification Workarounds  ------------------------------------
'------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
Public Function SheetIndexToName(SheetNum As Integer)
'    'FuncDefTest = RefCell.Address & ", " & RefCell.Address(True, True) & ", " & RefCell.Address(False, False) & ", " & RefCell.Address(True, False) & ", " & RefCell.Address(False, True) & ", " & Application.ThisCell.Address & ", " & Application.Caller.Parent.Name & ", " & Application.Caller.Parent.Name & ", " & ActiveCell.Row
'    Dim xRg As Range
'    Dim xRg1 As Range
'    Dim xCell As Range
'    Dim i As Long
'    Dim xAddress As String
'    Dim xUpdate As Boolean
'    Dim xRet As Variant
    SheetIndexToName = ThisWorkbook.Sheets(SheetNum).Name
End Function
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''------------------------------------------ SheetExists Function Development --------------------------------------
''------------------------------------------------------------------------------------------------------------------
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
' Resources:
'https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists'
''------------------------------------------------------------------------------------------------------------------
'******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''------------------------------------------------- Using Is Nothing T/F -------------------------------------------------
''------------------------------------------------------------------------------------------------------------------
''******************************************************************************************************************
''Public Function theWorksheetExists(ByVal theSheetName As String, Optional theWorkbook As Workbook) As Boolean
''    'Dim theWorkbook As Workbook
''    Dim testWorksheet As Worksheet
''    'Set theWorkbook = ActiveWorkbook
''    Set theWorkbook = ThisWorkbook
''    'Set testWorksheet = theWorkbook.
''    On Error Resume Next
''   theWorksheetExists = Not theWorkbook.Worksheets(theSheetName) Is Nothing
''End Function
''USAGE: If Not WorkSheetExists("SheetName") Then
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''------------------------------------------------- Using Set -------------------------------------------------
''------------------------------------------------------------------------------------------------------------------
''******************************************************************************************************************
'
''Set theWorksheet = Sheets(theSheetName)
''If theWorksheet Is Nothing Then
''     Sheets.Add(after:=ActiveSheet).Name = theSheetName
''     ''or whatever alternate code you want to execute''
''End If
''Function theWorksheetExists(theSheetName As String, Optional theWorkbook As Workbook) As Boolean
''    Dim theWorksheet As Worksheet
''
''     If theWorkbook Is Nothing Then Set theWorkbook = ThisWorkbook
''     'If theWorkbook Is Nothing Then Set theWorkbook = ActiveWorkbook
''     On Error Resume Next
''     Set theWorksheet = theWorkbook.Sheets(theSheetName)
''     On Error GoTo 0
''     theWorksheetExists = Not theWorksheet Is Nothing
'' End Function
''USAGE: If Contains(Sheets, "SheetName")
'Public Function theWorksheetExists(theSheetName As String, Optional theWorkbook As Workbook) As Boolean
'    If theWorkbook Is Nothing Then Set theWorkbook = ActiveWorkbook 'or ThisWorkbook - whichever appropriate
'    Dim theWorksheet As Object
'    On Error GoTo HandleError
'    Set theWorksheet = theWorkbook.Sheets(theSheetName)
'    theWorksheetExists = True
'    Exit Function
'HandleError:
'    theWorksheetExists = False
'End Function
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''------------------------------------------------- Loop Through Sheets -------------------------------------------------
''------------------------------------------------------------------------------------------------------------------
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
'For i = 1 To Worksheets.Count
'    If Worksheets(i).Name = theSheetName Then
'        theWorksheetExists = True
'    End If
'Next i
'
'If Not theWorksheetExists Then
'    Worksheets.Add.Name = theSheetName
'End If
''------------------------------------------------------------------------------------------------------------------
'Dim theWorkbook As Workbook
'Dim testWorksheet As Worksheet
''Set theWorkbook = ActiveWorkbook
'Set theWorkbook = ThisWorkbook
'For Each testWorksheet In theWorkbook.Worksheets
'    If testWorksheet.Name = theSheetName Then
'        'Do something here
'    End If
'Next
'Function theWorksheetExists(theSheetName As String) As Boolean
'    Dim theWorkbook As Workbook
'    Dim testWorksheet As Worksheet
'    'Set theWorkbook = ActiveWorkbook
'    Set theWorkbook = ThisWorkbook
'    theWorksheetExists = False
'      For Each testWorksheet In theWorkbook.Worksheets
'        If theSheetName = testWorksheet.Name Then
'          theWorksheetExists = True
'          Exit Function
'        End If
'      Next testWorksheet
'End Function
'Function theWorksheetExists(theSheetName As String) As Boolean
'    Dim i As Integer
'    For i = 1 To Worksheets.Count
'        If Worksheets(i).Name = theSheetName Then
'            theWorksheetExists = True
'            Exit Function
'        End If
'    Next i
'End Function
'Sub DoesTheSheetExists()
'    If SheetExist(theSheetName) Then
'        Debug.Print "The Sheet Exists"
'    Else
'        Debug.Print "The Sheet Does NOT Exists"
'    End If
'End Sub
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''------------------------------------------------- Using 'EVALUATE'/Error Checks -------------------------------------------------
''------------------------------------------------------------------------------------------------------------------
''******************************************************************************************************************
'Function theWorksheetExists(theSheetName As String) As Boolean
'    theWorksheetExists = Not WorksheetFunction.IsErr(Evaluate("'" & theSheetName & "'!A1"))
'End Function
'Function theWorksheetExists(theSheetName As String) As Boolean
'    theWorksheetExists = Evaluate("ISREF('" & theSheetName & "'!A1)")
'End Function
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''----------------------------------------------Generic Existence Check -------------------------------------------------
''------------------------------------------------------------------------------------------------------------------
''******************************************************************************************************************
'Function Contains(objCollection As Object, strName As String) As Boolean
'    Dim o As Object
'    On Error Resume Next
'    'Set o = objCollection(strName)
'    TypeName objCollection(strName)
'    Contains = (Err.Number = 0)
'    Err.Clear
' End Function
''******************************************************************************************************************
''------------------------------------------------------------------------------------------------------------------
''------------------------------------------------- Direct Cell Functions  -------------------------------------------------
''------------------------------------------------------------------------------------------------------------------
''******************************************************************************************************************
'Evaluate ("IsError(" & theSheetName & "!1:1)")
'WorksheetFunction.IsErr (Evaluate("'" & theSheetName & "'!A1"))
Sub SheetExists(theSheetName As String, Optional theWorkbook As Workbook)
    'If theWorkbook Is Nothing Then Set theWorkbook = ActiveWorkbook
    If theWorkbook Is Nothing Then Set theWorkbook = ThisWorkbook
    Dim theWorksheet As Worksheet
    Dim testWorksheet As Worksheet
    Dim theWorksheetObj As Object
    '------------------------------------------------------------------------------------------------------------------
    '<---- #1
    On Error Resume Next
    'Set theWorksheet = theWorkbook.Worksheets(theSheetName)             '+/-
    'On Error GoTo 0                                                     '+/-
    'SheetExists = Not theWorksheet Is Nothing                           '+/-
    SheetExists = Not theWorkbook.Worksheets(theSheetName) Is Nothing
'    '<---- #2
'    On Error GoTo HandleError
'    Set theWorksheetObj = theWorkbook.Sheets(theSheetName)
'    SheetExists = True
'    Exit Function
'HandleError:
'    SheetExists = False
'    '<---- #3
'    Set theWorksheet = theWorkbook.Sheets(theSheetName)
'    If theWorksheet Is Nothing Then
'         theWorkbook.Sheets.Add(after:=ActiveSheet).Name = theSheetName
'    End If
'    '<---- #4
'    Dim i As Integer
'    For i = 1 To theWorkbook.Worksheets.Count
'        If theWorkbook.Worksheets(i).Name = theSheetName Then
'            SheetExists = True
'            Exit Sub
'        End If
'    Next i
'    '<---- #5
'    SheetExists = False
'    For Each testWorksheet In theWorkbook.Worksheets
'        If theSheetName = testWorksheet.Name Then
'            SheetExists = True
'            Exit Sub
'        End If
'      Next testWorksheet
'    '<---- #6
'    SheetExists = Not WorksheetFunction.IsErr(Evaluate("'" & theSheetName & "'!A1"))
'    '<---- #7
'    SheetExists = Evaluate("ISREF('" & theSheetName & "'!A1)")
'    '<---- #8
'    SheetExists = Evaluate("IsError(" & theSheetName & "!1:1)")
End Sub
'********************************************************************************************************************************************
'--------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------- Notes ----------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------
'********************************************************************************************************************************************
