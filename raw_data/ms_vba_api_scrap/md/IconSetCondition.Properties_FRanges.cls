VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Private Const CLASS_NAME As String = "FRanges"
'----------------------------------------------------------------------------------------------------------
'Private mrBuild As Excel.Range
'----------------------------------------------------------------------------------------------------------


Public Function getNonEmptyRange(wks As Excel.Worksheet, _
                    Optional startFromFirstRow As Boolean = False, _
                    Optional startFromFirstColumn As Boolean = False) As Excel.Range
    Dim lastCell As Excel.Range
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim row As Long
    Dim col As Long
    Dim i As Long
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    'wks.AutoFilterMode = False
    
'[Doesn't work if sheet is protected]
'    Set lastCell = wks.Cells.SpecialCells(xlCellTypeLastCell)
'    If Not lastCell Is Nothing Then
'        With lastCell
'            lastRow = .row
'            lastCol = .column
'        End With
'    End If

    lastRow = lastNonEmptyRow(wks)
    lastCol = lastNonEmptyColumn(wks)
    
    

    
    If lastRow = 0 Or lastCol = 0 Then
        Set getNonEmptyRange = Nothing
    Else
        With wks
            arr = .Range(.cells(1, 1), .cells(lastRow, lastCol))
        End With
        
        If startFromFirstRow Then
            firstRow = 1
        Else
            For row = 1 To lastRow
                If Not VBA.IsEmpty(arr(row, lastCol)) Then
                    firstRow = row
                    Exit For
                End If
                col = wks.cells(row, lastCol).End(xlToLeft).column
                If col > 1 Then
                    firstRow = row
                    Exit For
                ElseIf Not VBA.IsEmpty(wks.cells(row, col)) Then
                    firstRow = row
                    Exit For
                End If
            Next row
        End If
        
        If startFromFirstColumn Then
            firstCol = 1
        Else
            For col = 1 To lastCol
                If Not VBA.IsEmpty(arr(lastRow, col)) Then
                    firstCol = col
                    Exit For
                End If
                row = wks.cells(lastRow, col).End(xlUp).row
                If row > 1 Then
                    firstCol = col
                    Exit For
                ElseIf Not VBA.IsEmpty(wks.cells(row, col)) Then
                    firstCol = col
                    Exit For
                End If
            Next col
        End If
        
        
        If firstRow <= lastRow And firstCol <= lastCol Then
            With wks
                Set getNonEmptyRange = .Range(.cells(firstRow, firstCol), .cells(lastRow, lastCol))
            End With
        End If
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
EmptySheetException:
    GoTo ExitPoint

End Function

Private Sub yieldBounds(rng As Excel.Range, ByRef firstRow As Long, ByRef firstCol As Long, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim subrng As Excel.Range
    Dim row As Long
    Dim col As Long
    '------------------------------------------------------------------------------------------------------
    
    If IsValid(rng) Then
    
        'First row
        For Each subrng In rng.Areas
            If firstRow = 1 Then Exit For
            row = subrng.row
            If row < firstRow Then firstRow = row
        Next subrng
        
        'First column
        For Each subrng In rng.Areas
            If firstCol = 1 Then Exit For
            col = subrng.column
            If col < firstCol Then firstCol = col
        Next subrng
        
        'Last row
        For Each subrng In rng.Areas
            row = subrng.row + subrng.Rows.count - 1
            If row > lastRow Then lastRow = row
        Next subrng
        
        'Last column
        For Each subrng In rng.Areas
            col = subrng.column + subrng.columns.count - 1
            If col > lastCol Then lastCol = col
        Next subrng
        
    End If
    
End Sub

Public Function getNonEmptyCells(rng As Excel.Range) As Excel.Range
    Dim rngFormulas As Variant 'Excel.Range
    Dim rngConstants As Variant 'Excel.Range
    '----------------------------------------------------------------------

    On Error Resume Next
    With rng
        Set rngFormulas = .SpecialCells(xlCellTypeFormulas)
        Set rngConstants = .SpecialCells(xlCellTypeConstants)
    End With
    On Error GoTo 0

    Set getNonEmptyCells = f.ranges.join(f.Collections.Create(rngFormulas, rngConstants))

End Function

Public Function join(ParamArray ranges() As Variant) As Excel.Range
    Dim arrRanges As Variant:   arrRanges = ranges
    Dim colRanges As VBA.Collection
    Dim varRange As Variant
    Dim rng As Excel.Range
    Dim Parent As Excel.Worksheet
    '------------------------------------------------------------------------------------------------------

    Set colRanges = New VBA.Collection

    For Each varRange In ranges
        If VBA.TypeName(varRange) = "Range" Then
            Call colRanges.Add(varRange)
        ElseIf f.Variables.isContainer(varRange) Then
            Set colRanges = f.Collections.join(colRanges, varRange)
        End If
    Next varRange

    For Each varRange In colRanges
        If VBA.IsObject(varRange) Then
            If varRange Is Nothing Then
            
            ElseIf TypeOf varRange Is Excel.Range Then
                Set rng = varRange
                If IsValid(rng) Then
                    If join Is Nothing Then
                        Set join = rng
                        Set Parent = rng.Parent
                    Else
                        If Not rng.Parent Is Parent Then Call VBA.err.raise(1, description:="DifferentWorksheetsException")
                        Set rng = varRange
                        DoEvents
                        Set join = rng.Application.Union(join, rng)
                    End If
                End If
            End If
        End If
    Next varRange
    
End Function



'Public Function getRangeFirstColumn(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    '------------------------------------------------------------------------------------------------------
'
'    Call throwExceptionIfRangeIsInvalid(rng)
'    getRangeFirstColumn = rng.parent.columns.Count + 1
'    For Each subrange In rng.Areas
'        getRangeFirstColumn = Excel.Application.WorksheetFunction.Min(getRangeFirstColumn, subrange.column)
'    Next subrange
'
'End Function
'
'Public Function getRangeFirstRow(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    '------------------------------------------------------------------------------------------------------
'
'    If isRangeValid(rng) Then
'        getRangeFirstRow = rng.parent.rows.Count + 1
'        For Each subrange In rng.Areas
'            getRangeFirstRow = Excel.Application.WorksheetFunction.Min(getRangeFirstRow, subrange.row)
'        Next subrange
'    End If
'
'End Function
'
'Public Function getRangeLastColumn(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    Dim subrangeLastColumn As Long
'    '------------------------------------------------------------------------------------------------------
'
'    If isRangeValid(rng) Then
'        For Each subrange In rng.Areas
'            subrangeLastColumn = subrange.column + subrange.columns.Count - 1
'            getRangeLastColumn = Excel.Application.WorksheetFunction.Max(getRangeLastColumn, subrangeLastColumn)
'        Next subrange
'    End If
'
'End Function
'
'Public Function getRangeLastRow(rng As Excel.Range) As Long
'    Dim subrange As Excel.Range
'    Dim subrangeLastRow As Long
'    '------------------------------------------------------------------------------------------------------
'
'    If isRangeValid(rng) Then
'        For Each subrange In rng.Areas
'            subrangeLastRow = subrange.row + subrange.rows.Count - 1
'            getRangeLastRow = Excel.Application.WorksheetFunction.Max(getRangeLastRow, subrangeLastRow)
'        Next subrange
'    End If
'
'End Function
'


Public Function findCellsByContent(rng As Range, what As Variant, Optional lookIn As XlFindLookIn = xlValues, Optional lookAt As XlLookAt = xlWhole, _
                        Optional searchOrder As XlSearchOrder = xlByColumns, Optional searchDirection As XlSearchDirection = xlNext, _
                        Optional matchCase As Boolean = False, Optional matchByte As Boolean = False, Optional searchFormat As Boolean = False) As Excel.Range
    Dim searchResult As Range
    Dim firstMatch As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    With rng
        Set searchResult = .Find(what, , lookIn, lookAt, searchOrder, searchDirection, matchCase, matchByte, searchFormat)
        If Not searchResult Is Nothing Then
            Set firstMatch = searchResult
            'firstMatch = searchResult.address
            
            Do
                If findCellsByContent Is Nothing Then
                    Set findCellsByContent = searchResult
                Else
                    Set findCellsByContent = join(findCellsByContent, searchResult)
                End If
                Set searchResult = .FindNext(searchResult)
                
                If searchResult Is Nothing Then
                    Exit Do
                ElseIf Not rng.Application.intersect(searchResult, firstMatch) Is Nothing Then
                    Exit Do
                End If
            Loop
            
        End If
    End With
    
End Function

Public Function findRowByLabel(wks As Excel.Worksheet, label As String) As Long
    Dim cell As Excel.Range
    '------------------------------------------------------------------------------------------------------
    With wks
        Set cell = .cells.Find(what:=label, After:=.cells(1, 1), lookAt:=xlWhole, lookIn:=xlFormulas, searchOrder:=xlByRows, searchDirection:=xlNext, matchCase:=False)
        If Not cell Is Nothing Then findRowByLabel = cell.row
    End With
End Function

Public Function findColumnByLabel(wks As Excel.Worksheet, label As String) As Long
    Dim cell As Excel.Range
    '------------------------------------------------------------------------------------------------------
    With wks
        Set cell = .cells.Find(what:=label, After:=.cells(1, 1), lookAt:=xlWhole, lookIn:=xlFormulas, searchOrder:=xlByRows, searchDirection:=xlNext, matchCase:=False)
        If Not cell Is Nothing Then findColumnByLabel = cell.column
    End With
End Function

Public Sub setInsideBorders(rng As Excel.Range, Optional color As Long = 14277081)
    With rng.Borders(xlInsideHorizontal)
        .color = color
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .color = color
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Public Sub clearInsideBorders(rng As Excel.Range)
    If f.ranges.IsValid(rng) Then
        With rng.Borders(xlInsideHorizontal)
            .LineStyle = xlLineStyleNone
        End With
        With rng.Borders(xlInsideVertical)
            .LineStyle = xlLineStyleNone
        End With
    End If
End Sub



'If the given range is a few cells merged in one area - return only the first of them.
Public Function getFirstCell(rng As Excel.Range) As Excel.Range
    Dim objName As Excel.name
    Dim firstCell As Excel.Range
    Dim mergedArea As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If rng.cells.count > 1 Then
        Set firstCell = rng.cells(1, 1)
        If firstCell.MergeCells Then
            Set mergedArea = firstCell.MergeArea
            If f.strings.compareStrings(mergedArea.address, rng.address) Then
                Set getFirstCell = firstCell
            End If
        End If
    End If
    
    If getFirstCell Is Nothing Then
        Set getFirstCell = rng
    End If
    
End Function

Public Function IsValid(rng As Excel.Range) As Boolean
    Dim lngRangeRow As Long
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    lngRangeRow = rng.row
    If lngRangeRow > 0 Then IsValid = True
End Function

Public Function pasteData(data As Variant, initRange As Excel.Range, Optional transponeData As Boolean = True, _
                        Optional clearSheet As Boolean = False, Optional sheetPassword As Variant) As Excel.Range
    Dim arr As Variant
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    Dim enableEventsCache As Boolean
    Dim protected As Boolean
    Dim filter As ExcelFilter
    '------------------------------------------------------------------------------------------------------
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    '------------------------------------------------------------------------------------------------------

    
    With Application
        enableEventsCache = .EnableEvents
        .EnableEvents = False
    End With
    
    If Not IsValid(initRange) Then Call Exceptions.InvalidRange.raise
    
    arr = f.arrays.to2DArray(data)
    If f.arrays.countDimensions(arr) <> 2 Then Call Exceptions.InvalidDimension.raise
    
    With initRange
        
        Set wks = .Parent
        
        If wks.ProtectContents And Not VBA.IsMissing(sheetPassword) Then
            protected = True
            Call wks.Unprotect(sheetPassword)
        End If
        
        If clearSheet Then
            Call wks.cells.ClearContents
        Else
            Set filter = f.Create.ExcelFilter(wks)
            Call filter.readFromSheet
        End If
        
        If transponeData Then arr = f.arrays.transposeArray(arr)
                                                                                                        
        firstRow = .row
        firstCol = .column
        lastRow = .row + f.arrays.getSize(arr, 1) - 1
        lastCol = .column + f.arrays.getSize(arr, 2) - 1
                                                                                                        
    End With
        
    With wks
        Set rng = .Range(.cells(firstRow, firstCol), .cells(lastRow, lastCol))
    End With
    
    If Not filter Is Nothing Then Call filter.removeFilter
    rng = arr
    If Not filter Is Nothing Then Call filter.applyFilter
    
    Set pasteData = rng

'==========================================================================================================
ExitPoint:
    
    Application.EnableEvents = enableEventsCache
    If protected Then Call wks.Protect(sheetPassword)
    
    Exit Function

End Function

Public Function getCellAddress(cell As Excel.Range, Optional rowLocked As Boolean = True, Optional columnLocked As Boolean = True)
    Dim defaultAddress As String
    Dim columnIndex As String
    Dim row As Long
    '------------------------------------------------------------------------------------------------------
    defaultAddress = cell.address
    columnIndex = f.strings.substring(defaultAddress, "$", "$")
    row = VBA.CLng(f.strings.substring(defaultAddress, columnIndex & "$", vbNullString))
    getCellAddress = VBA.IIf(columnLocked, "$", vbNullString) & columnIndex & VBA.IIf(rowLocked, "$", vbNullString) & row
End Function

Public Sub clearFormatConditions(rng As Excel.Range)
    Dim fc As Excel.FormatCondition
    '------------------------------------------------------------------------------------------------------
    If IsValid(rng) Then
        For Each fc In rng.FormatConditions
            Call fc.delete
        Next fc
    End If
End Sub

Public Sub pasteAsValues(rng As Excel.Range, Optional protectionPassword As String)
    Dim wks As Excel.Worksheet
    Dim currentSelection As Excel.Range
    Dim isProtected As Boolean
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    Set currentSelection = Excel.selection
    On Error GoTo 0
    
    If IsValid(rng) Then
        With rng
            Set wks = .Parent
            isProtected = wks.ProtectContents
            If isProtected Then Call wks.Unprotect(protectionPassword)
            Call .copy
            Call .PasteSpecial(Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False)
            If isProtected Then Call wks.Protect(protectionPassword)
        End With
    End If

    On Error Resume Next
    Application.CutCopyMode = False
    Call currentSelection.Select

End Sub





Public Function getFirstNonEmptyRow(wks As Excel.Worksheet, _
                      Optional startRow As Long, Optional startCol As Long, _
                      Optional endRow As Long, Optional endCol As Long, _
                      Optional ignoreHiddenCells As Boolean = False) As Long
    Const METHOD_NAME As String = "firstNonEmptyRow"
    '------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim lngNonBlanks As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------


    'Checks if the given worksheet is valid and can be referred to. -------------------------------------|
    'If not, the code jumps to the label IllegalSheetException, where you can define your own           '|
    'error handling rules for this exception.                                                           '|
    If Not f.sheets.IsValid(wks) Then GoTo IllegalSheetException                                             '|
    '----------------------------------------------------------------------------------------------------|


    'Calculate the actual range to be searched including optional parameters passed to the function. ----|
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1  '|
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count   '|
    If startRow > 0 And startRow <= wks.Rows.count Then lngRowStart = startRow Else lngRowStart = 1     '|
    If endRow > 0 And endRow <= wks.Rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.Rows.count  '|
    '----------------------------------------------------------------------------------------------------|




Retry:

    '----------------------------------------------------------------------------------------------------|
    If endRow > 0 And endRow <= wks.Rows.count Then                                                     '|
        lngRow = endRow                                                                                 '|
    Else                                                                                                '|
        lngRow = wks.Rows.count                                                                         '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|


    '----------------------------------------------------------------------------------------------------|
    Do                                                                                                  '|
        Set rng = wks.Range(wks.cells(lngRowStart, lngColStart), wks.cells(lngRow, lngColEnd))          '|
        lngNonBlanks = wks.Application.WorksheetFunction.CountA(rng)                                    '|
                                                                                                        '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If lngNonBlanks Then                                                                        '|  '|
                                                                                                    '|  '|
            If lngRow = lngRowStart Then Exit Do                                                    '|  '|
            lngRowEnd = lngRow                                                                      '|  '|
            lngRow = lngRowStart + ((lngRow - lngRowStart - 1) / 2)                                 '|  '|
                                                                                                    '|  '|
        Else                                                                                        '|  '|
            lngRowStart = lngRow + 1                                                                '|  '|
            lngRow = lngRowEnd                                                                      '|  '|
                                                                                                    '|  '|
            '------------------------------------------------------------------------------------|  '|  '|
            If lngRowStart > lngRowEnd Then                                                     '|  '|  '|
                lngRow = 0                                                                      '|  '|  '|
                Exit Do                                                                         '|  '|  '|
            End If                                                                              '|  '|  '|
            '------------ [If lngRowStart > lngRowEnd Then] -------------------------------------|  '|  '|
                                                                                                    '|  '|
        End If                                                                                      '|  '|
        '---------------- [If lngNonBlanks Then] ----------------------------------------------------|  '|
                                                                                                        '|
    Loop                                                                                                '|
    '----------------------------------------------------------------------------------------------------|


    '----------------------------------------------------------------------------------------------------|
    If lngRow Then                                                                                      '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If ignoreHiddenCells And wks.Rows(lngRow).Hidden Then                                       '|  '|
            lngRowStart = nextVisibleRow(wks, lngRow, Excel.xlDown)                                 '|  '|
            If lngRowStart <= wks.Rows.count Then GoTo Retry                                        '|  '|
        Else                                                                                        '|  '|
            getFirstNonEmptyRow = lngRow                                                            '|  '|
        End If                                                                                      '|  '|
        '---------------- [If ignoreHiddenCells And wks.rows(lngRow).Hidden Then] -------------------|  '|
                                                                                                        '|
    End If                                                                                              '|




'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function firstNonEmptyColumn(wks As Excel.Worksheet, _
                   Optional startRow As Long, Optional startCol As Long, _
                   Optional endRow As Long, Optional endCol As Long, _
                   Optional ignoreHiddenCells As Boolean = False) As Long
    Const METHOD_NAME As String = "firstNonEmptyColumn"
    '------------------------------------------------------------------------------------------------------
    Dim lngCol As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim lngNonBlanks As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------


    'Checks if the given worksheet is valid and can be referred to. -------------------------------------|
    'If not, the code jumps to the label IllegalSheetException, where you can define your own           '|
    'error handling rules for this exception.                                                           '|
    If Not f.sheets.IsValid(wks) Then GoTo IllegalSheetException                                            '|
    '----------------------------------------------------------------------------------------------------|
    
    
    'Calculate the actual range to be searched including optional parameters passed to the function. ----|
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1  '|
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count   '|
    If startRow > 0 And startRow <= wks.Rows.count Then lngRowStart = startRow Else lngRowStart = 1     '|
    If endRow > 0 And endRow <= wks.Rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.Rows.count  '|
    '----------------------------------------------------------------------------------------------------|


Retry:

    '----------------------------------------------------------------------------------------------------|
    If endCol > 0 And endCol <= wks.columns.count Then                                                  '|
        lngCol = endCol                                                                                 '|
    Else                                                                                                '|
        lngCol = wks.columns.count                                                                      '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|
    

    '----------------------------------------------------------------------------------------------------|
    Do                                                                                                  '|
                                                                                                        '|
        Set rng = wks.Range(wks.cells(lngRowStart, lngColStart), wks.cells(lngRowEnd, lngCol))          '|
        lngNonBlanks = wks.Application.WorksheetFunction.CountA(rng)                                    '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If lngNonBlanks Then                                                                        '|  '|
                                                                                                    '|  '|
            If lngCol = lngColStart Then Exit Do                                                    '|  '|
            lngColEnd = lngCol                                                                      '|  '|
            lngCol = lngColStart + ((lngCol - lngColStart - 1) / 2)                                 '|  '|
                                                                                                    '|  '|
        Else                                                                                        '|  '|
                                                                                                    '|  '|
            lngColStart = lngCol + 1                                                                '|  '|
            lngCol = lngColEnd                                                                      '|  '|
                                                                                                    '|  '|
            '------------------------------------------------------------------------------------|  '|  '|
            If lngColStart > lngColEnd Then                                                     '|  '|  '|
                lngCol = 0                                                                      '|  '|  '|
                Exit Do                                                                         '|  '|  '|
            End If                                                                              '|  '|  '|
            '------ [If lngColStart > lngColEnd Then] -------------------------------------------|  '|  '|
                                                                                                    '|  '|
        End If                                                                                      '|  '|
        '---------- [If lngNonBlanks Then] ----------------------------------------------------------|  '|
                                                                                                        '|
    Loop                                                                                                '|
    '----------------------------------------------------------------------------------------------------|
        

    '----------------------------------------------------------------------------------------------------|
    If lngCol Then                                                                                      '|
                                                                                                        '|
        '--------------------------------------------------------------------------------------------|  '|
        If ignoreHiddenCells And wks.columns(lngCol).Hidden Then                                    '|  '|
            lngColStart = nextVisibleColumn(wks, lngCol, Excel.xlToRight)                           '|  '|
            If lngColStart <= wks.columns.count Then GoTo Retry                                     '|  '|
        Else                                                                                        '|  '|
            firstNonEmptyColumn = lngCol                                                            '|  '|
        End If                                                                                      '|  '|
        '--------------------------------------------------------------------------------------------|  '|
                                                                                                        '|
    End If                                                                                              '|
    '----------------------------------------------------------------------------------------------------|



'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function


Public Function getLastNonEmptyRow(wks As Excel.Worksheet, _
                      Optional startRow As Long, Optional startCol As Long, _
                      Optional endRow As Long, Optional endCol As Long, _
                      Optional ignoreHiddenCells As Boolean = False) As Long
    getLastNonEmptyRow = lastNonEmptyRow(wks, startRow, startCol, endRow, endCol, ignoreHiddenCells)
End Function

Public Function lastNonEmptyRow(wks As Excel.Worksheet, _
                      Optional startRow As Long, Optional startCol As Long, _
                      Optional endRow As Long, Optional endCol As Long, _
                      Optional ignoreHiddenCells As Boolean = False) As Long
    Dim lngRow As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim lngNonBlanks As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------

    If Not f.sheets.IsValid(wks) Then GoTo IllegalSheetException
    
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count
    If startRow > 0 And startRow <= wks.Rows.count Then lngRowStart = startRow Else lngRowStart = 1
    If endRow > 0 And endRow <= wks.Rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.Rows.count

Retry:

    lngRow = 1

    Do
        Set rng = wks.Range(wks.cells(lngRow, lngColStart), wks.cells(lngRowEnd, lngColEnd))
        lngNonBlanks = wks.Application.WorksheetFunction.CountA(rng)
                                                                                                        
        If lngNonBlanks Then
            If lngRow = lngRowEnd Then Exit Do
            lngRowStart = lngRow
            lngRow = lngRow + ((lngRowEnd - lngRow + 1) / 2)
        Else
            lngRowEnd = lngRow - 1
            lngRow = lngRowStart
                                                                                                      
            If lngRowStart > lngRowEnd Then
                lngRow = 0
                Exit Do
            End If
                                                                                                      
        End If
    Loop

    If lngRow Then
        If ignoreHiddenCells And wks.Rows(lngRow).Hidden Then
            lngRowEnd = nextVisibleRow(wks, lngRow, Excel.xlUp)
            If lngRowEnd Then GoTo Retry
        Else
            lastNonEmptyRow = lngRow
        End If
    End If

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function getLastNonEmptyColumn(wks As Excel.Worksheet, _
                   Optional startRow As Long, Optional startCol As Long, _
                   Optional endRow As Long, Optional endCol As Long, _
                   Optional ignoreHiddenCells As Boolean = False) As Long
    getLastNonEmptyColumn = lastNonEmptyColumn(wks, startRow, startCol, endRow, endCol, ignoreHiddenCells)
End Function

Public Function lastNonEmptyColumn(wks As Excel.Worksheet, _
                   Optional startRow As Long, Optional startCol As Long, _
                   Optional endRow As Long, Optional endCol As Long, _
                   Optional ignoreHiddenCells As Boolean = False) As Long
    Const METHOD_NAME As String = "lastNonEmptyColumn"
    '------------------------------------------------------------------------------------------------------
    Dim xls As Excel.Application
    Dim lngCol As Long
    Dim lngColStart As Long
    Dim lngColEnd As Long
    Dim lngNonBlanks As Long
    Dim lngRowStart As Long
    Dim lngRowEnd As Long
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------

    If Not f.sheets.IsValid(wks) Then GoTo IllegalSheetException
    Set xls = wks.Parent.Application
    
    If startCol > 0 And startCol <= wks.columns.count Then lngColStart = startCol Else lngColStart = 1
    If endCol > 0 And endCol <= wks.columns.count Then lngColEnd = endCol Else _
                                                                        lngColEnd = wks.columns.count
    If startRow > 0 And startRow <= wks.Rows.count Then lngRowStart = startRow Else lngRowStart = 1
    If endRow > 0 And endRow <= wks.Rows.count Then lngRowEnd = endRow Else lngRowEnd = wks.Rows.count

Retry:
    lngCol = 1
    
    Do
        Set rng = wks.Range(wks.cells(lngRowStart, lngCol), wks.cells(lngRowEnd, lngColEnd))
        lngNonBlanks = xls.WorksheetFunction.CountA(rng)
                                                                                                        
        If lngNonBlanks Then
                                                                                                      
            If lngCol = lngColEnd Then Exit Do
            lngColStart = lngCol
            lngCol = lngCol + ((lngColEnd - lngCol + 1) / 2)
                                                                                                      
        Else
            lngColEnd = lngCol - 1
            lngCol = lngColStart
                                                                                                      
            If lngColStart > lngColEnd Then
                lngCol = 0
                Exit Do
            End If
                                                                                                      
        End If
                                                                                                        
    Loop

    If lngCol Then
        If ignoreHiddenCells And wks.columns(lngCol).Hidden Then
            lngColEnd = nextVisibleColumn(wks, lngCol, Excel.xlToLeft)
            If lngColEnd Then GoTo Retry
        Else
            lastNonEmptyColumn = lngCol
        End If
    End If


'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function nextVisibleRow(wks As Excel.Worksheet, initialRow As Long, _
                                                                          direction As XlDirection) As Long
    Dim intOffset As Integer
    '------------------------------------------------------------------------------------------------------
                
    If Not f.sheets.IsValid(wks) Then Call Exceptions.IllegalSheet.raise
                
    Select Case direction
        Case Excel.xlUp:    intOffset = -1
        Case Excel.xlDown:  intOffset = 1
        Case Else
            nextVisibleRow = initialRow
            GoTo ExitPoint
    End Select

    nextVisibleRow = initialRow
    Do
        nextVisibleRow = nextVisibleRow + intOffset
        If Not wks.Rows(nextVisibleRow).Hidden Then Exit Do
    Loop

'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function

Public Function nextVisibleColumn(wks As Excel.Worksheet, initialCol As Long, _
                                          direction As XlDirection) As Long
    Dim intOffset As Integer
    '------------------------------------------------------------------------------------------------------
                
    If Not f.sheets.IsValid(wks) Then Call Exceptions.IllegalSheet.raise
                
    Select Case direction
        Case Excel.xlToLeft:    intOffset = -1
        Case Excel.xlToRight:   intOffset = 1
        Case Else
            nextVisibleColumn = initialCol
            GoTo ExitPoint
    End Select


    nextVisibleColumn = initialCol
    Do
        nextVisibleColumn = nextVisibleColumn + intOffset
        If Not wks.columns(nextVisibleColumn).Hidden Then Exit Do
    Loop


'==========================================================================================================
ExitPoint:
    Exit Function

    '----------------------------------------------------------------------------------------------------------
IllegalSheetException:
    GoTo ExitPoint

End Function



Public Function intersect(ParamArray ranges() As Variant) As Excel.Range
    Dim arrRanges As Variant:               arrRanges = ranges
    Dim varRange As Variant
    Dim i As Integer
    Dim isInitialized As Boolean
    Dim rng As Excel.Range
    Dim subrng As Excel.Range
    Dim app As Excel.Application
    '------------------------------------------------------------------------------------------------------

    If f.arrays.isDefined(arrRanges) Then
        For i = LBound(arrRanges) To UBound(arrRanges)
            Call f.Variables.assign(varRange, arrRanges(i))

            If isNotEmptyRange(varRange) Then
                If Not isInitialized Then
                    Set rng = varRange
                    Set app = rng.Application
                    isInitialized = True
                Else
                    Set rng = app.intersect(rng, varRange)
                    If rng Is Nothing Then Exit For
                End If
            Else
                Set rng = Nothing
                Exit For
            End If

        Next i
    End If

    Set intersect = rng

End Function

Private Function isNotEmptyRange(rng As Variant) As Boolean
    Const METHOD_NAME As String = "isNotEmptyRange"
    '------------------------------------------------------------------------------------------------------
    Dim subrng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    If VBA.IsObject(rng) Then
        If Not rng Is Nothing Then
            If TypeOf rng Is Excel.Range Then
                Set subrng = rng
                isNotEmptyRange = f.ranges.IsValid(subrng)
            End If
        End If
    End If
End Function


Public Function getNextEmptyCell(cell As Excel.Range, direction As Long) As Excel.Range
    Dim wks As Excel.Worksheet
    Dim i As Long
    Dim stepSize As Long
    Dim firstIndex As Long
    Dim lastIndex As Long
    Dim isHorizontalSearch As Boolean
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(cell) Then
        Set wks = cell.Parent
        
        If direction = xlToLeft Or direction = xlUp Then
            stepSize = -1
            lastIndex = 1
        ElseIf direction = xlToRight Or direction = xlDown Then
            stepSize = 1
            lastIndex = VBA.IIf(direction = xlToRight, wks.columns.count, wks.Rows.count)
        End If
        
        isHorizontalSearch = (direction = xlToRight Or direction = xlToLeft)
        firstIndex = VBA.IIf(isHorizontalSearch, cell.column, cell.row)
        
        For i = firstIndex To lastIndex Step stepSize
            If isHorizontalSearch Then
                Set rng = wks.cells(cell.row, i)
            Else
                Set rng = wks.cells(i, cell.column)
            End If
            If VBA.IsEmpty(rng.value) Then
                Set getNextEmptyCell = rng
                Exit For
            End If
        Next i
        
    End If
    
End Function



Public Function insertColumns(startCell As Excel.Range, columnsCounter As Long) As Excel.Range
    Dim wks As Excel.Worksheet
    Dim rng As Excel.Range
    Dim column As Long
    '------------------------------------------------------------------------------------------------------
    
    Set wks = startCell.Parent
    column = startCell.column
    
    Set rng = startCell.Resize(1, columnsCounter).EntireColumn
    Call rng.insert(Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove)
    
    With wks
        Set insertColumns = .Range(.cells(1, column), .cells(1, column + columnsCounter - 1)).EntireColumn
    End With
    
End Function

Public Function insertRows(Sheet As Excel.Worksheet, topRow As Long, rowsCounter As Long, _
                           Optional Shift As Variant = xlDown, Optional clearFormat As Boolean = True) As Excel.Range
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    With Sheet
        Set rng = .Range(.cells(topRow, 1), .cells(topRow + rowsCounter - 1, 1)).EntireRow
        Call rng.insert(VBA.IIf(Shift = xlDown, -4121, -4162))
        Set rng = .Range(.cells(topRow, 1), .cells(topRow + rowsCounter - 1, 1)).EntireRow
        If clearFormat Then Call rng.clear
    End With
    Set insertRows = rng
End Function


Public Function getRangeLastColumn(rng As Excel.Range) As Long
    Dim subrng As Excel.Range
    Dim lastCol As Long
    '------------------------------------------------------------------------------------------------------
    On Error Resume Next
    If f.ranges.IsValid(rng) Then
        For Each subrng In rng.Areas
            With subrng
                lastCol = .column + .columns.count - 1
                If lastCol > getRangeLastColumn Then getRangeLastColumn = lastCol
            End With
        Next subrng
    End If
End Function

Public Function getRangeLastRow(rng As Excel.Range) As Long
    Dim subrng As Excel.Range
    Dim lastRow As Long
    '------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    If f.ranges.IsValid(rng) Then
        For Each subrng In rng.Areas
            With subrng
                lastRow = .row + .Rows.count - 1
                If lastRow > getRangeLastRow Then getRangeLastRow = lastRow
            End With
        Next subrng
    End If
    
End Function







'borderDict - JSON object width definition of borders.
' Template:
'            {
'                "right": {
'                    "color": 16777215,
'                    "weight": "thin",
'                    "style": "continuous"
'                },
'                "left": {...},
'                "top": {...},
'                "bottom": {...},
'                "inside-horizontal": {...},
'                "inside-vertical": {...}
'            }
Public Sub applyBorders(rng As Excel.Range, ByVal bordersDict As Scripting.Dictionary, _
                            Optional skipOutsideBorders As Boolean = False, Optional skipInsideBorders As Boolean = False)
    Dim varKey As Variant
    Dim borderIndex As XlBordersIndex
    Dim borderProps As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    Set bordersDict = getBordersDefinitionsFromDict( _
                            f.dictionaries.createSingleEntryDict(UIProps.Borders.getName, bordersDict))
    
    If f.dictionaries.isNonEmptyDictionary(bordersDict) Then
        For Each varKey In bordersDict.Keys
            borderIndex = convertBorderIndexNameToEnum(VBA.CStr(varKey))
            If (isInsideBorder(borderIndex) And Not skipInsideBorders) Or (isOutsideBorder(borderIndex) And Not skipOutsideBorders) Then
                Set borderProps = f.dictionaries.getObject(bordersDict, varKey)
                If Not borderProps Is Nothing Then Call applyBorder(rng, borderIndex, borderProps)
            End If
        Next varKey
    End If
    
End Sub

Public Sub applySpecificBorder(rng As Excel.Range, bordersDict As Scripting.Dictionary, borderIndex As XlBordersIndex)
    Dim borderIndexName As String
    Dim borderProps As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If f.dictionaries.isNonEmptyDictionary(bordersDict) Then
        borderIndexName = convertBorderIndexEnumToName(borderIndex)
        Set borderProps = f.dictionaries.getObject(bordersDict, borderIndexName)
        If Not borderProps Is Nothing Then
            Call applyBorder(rng, borderIndex, borderProps)
        Else
            rng.Borders(borderIndex).LineStyle = xlLineStyleNone
        End If
    End If
    
End Sub

Private Sub applyBorder(rng As Excel.Range, borderIndex As XlBordersIndex, borderProps As Scripting.Dictionary)
    Const BORDER_WEIGHT_TAG As String = "weight"
    Const BORDER_COLOR_TAG As String = "color"
    Const BORDER_STYLE_TAG As String = "style"
    '------------------------------------------------------------------------------------------------------
    With rng.Borders(borderIndex)
        .Weight = convertBorderWeightNameToEnum(f.dictionaries.getItem(borderProps, BORDER_WEIGHT_TAG))
        .LineStyle = convertBorderStyleNameToEnum(f.dictionaries.getItem(borderProps, BORDER_STYLE_TAG))
        .color = f.dictionaries.getItem(borderProps, BORDER_COLOR_TAG)
    End With
End Sub

Private Sub applyAdjacentBorder(rng As Excel.Range, borderIndex As XlBordersIndex, borderProps As Scripting.Dictionary)
    Dim rngAdjacent As Excel.Range
    Dim oppositeBorderIndex As XlBordersIndex
    '------------------------------------------------------------------------------------------------------
    
    If borderIndex = xlEdgeBottom Then
        oppositeBorderIndex = xlEdgeTop
        Set rngAdjacent = rng.Rows(rng.Rows.count).offset(1)
    ElseIf borderIndex = xlEdgeTop Then
        oppositeBorderIndex = xlEdgeBottom
        Set rngAdjacent = rng.Rows(1).offset(-1)
    ElseIf borderIndex = xlEdgeLeft Then
        oppositeBorderIndex = xlEdgeRight
        Set rngAdjacent = rng.columns(1).offset(, -1)
    ElseIf borderIndex = xlEdgeRight Then
        oppositeBorderIndex = xlEdgeLeft
        Set rngAdjacent = rng.columns(rng.columns.count).offset(, 1)
    End If
    
    If f.ranges.IsValid(rngAdjacent) And oppositeBorderIndex > 0 Then
        Call applyBorder(rngAdjacent, oppositeBorderIndex, borderProps)
    End If
    
End Sub





Public Function applyFormatCondition(rng As Excel.Range, formatProps As Scripting.Dictionary) As Object
    Dim formatType As XlFormatConditionType
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        formatType = f.Reflection.modifyValue(f.dictionaries.getItem(formatProps, "type"), _
                            f.Reflection.getFunctionFullName("convertFormatConditionTypeToEnum", Excel.ThisWorkbook))
        Select Case formatType
            Case xlIconSets:                    Set applyFormatCondition = applyIconSetsFormatCondition(rng, formatProps)
            Case xlExpression, xlCellValue:     Set applyFormatCondition = applyRegularFormatCondition(rng, formatProps)
            Case Else:                          Set applyFormatCondition = applyRegularFormatCondition(rng, formatProps)
        End Select
    End If
    
End Function

Private Function applyRegularFormatCondition(rng As Excel.Range, formatProps As Scripting.Dictionary) As Excel.FormatCondition
    Dim formatType As XlFormatConditionType
    Dim operator As XlFormatConditionOperator
    Dim setFirstPriority As Boolean
    Dim priority As Variant
    Dim stopIfTrue As Boolean
    Dim isFormulaR1C1 As Boolean
    Dim formula As Variant
    Dim backColor As Variant
    Dim fontColor As Variant
    Dim fontSize As Variant
    Dim fontBold As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        formatType = f.Variables.coalesce( _
                        f.Reflection.modifyValue(f.dictionaries.getItem(formatProps, "type"), _
                            f.Reflection.getFunctionFullName("convertFormatConditionTypeToEnum", Excel.ThisWorkbook)), _
                     xlExpression)
        operator = f.Variables.coalesce( _
                        f.Reflection.modifyValue(f.dictionaries.getItem(formatProps, "operator"), _
                            f.Reflection.getFunctionFullName("convertFormatConditionOperatorToEnum", Excel.ThisWorkbook)), _
                     xlEqual)
        setFirstPriority = f.dictionaries.getItem(formatProps, "setFirstPriority")
        priority = f.dictionaries.getItem(formatProps, "priority")
        stopIfTrue = f.dictionaries.getItem(formatProps, "stopIfTrue")
        isFormulaR1C1 = VBA.IIf(formatProps.exists(props.isFormulaR1C1.getName), _
                                f.dictionaries.getItem(formatProps, props.isFormulaR1C1.getName), _
                                True)
        formula = f.Reflection.convertToLocalFormula(f.dictionaries.getItem(formatProps, "formula"), isFormulaR1C1)
        backColor = f.dictionaries.getItemForCombinedKey(formatProps, "interior.color")
        fontColor = f.dictionaries.getItemForCombinedKey(formatProps, "font.color")
        fontSize = f.dictionaries.getItemForCombinedKey(formatProps, "font.size")
        fontBold = f.dictionaries.getItemForCombinedKey(formatProps, "font.bold")
        
        '[Apply properties to range]
        Set applyRegularFormatCondition = rng.FormatConditions.Add(formatType, operator, formula)
        With applyRegularFormatCondition
            If setFirstPriority Then Call .setFirstPriority
            .stopIfTrue = stopIfTrue
            If Not VBA.IsEmpty(priority) Then .priority = priority
            If Not VBA.IsEmpty(backColor) Then .Interior.color = backColor
            If Not VBA.IsEmpty(fontColor) Then .Font.color = fontColor
            If Not VBA.IsEmpty(fontBold) Then .Font.Bold = fontBold
            If Not VBA.IsEmpty(fontSize) Then .Font.size = fontSize
        End With
    End If
    
End Function

Private Function applyIconSetsFormatCondition(rng As Excel.Range, formatProps As Scripting.Dictionary) As Excel.IconSetCondition
    Dim stopIfTrue As Boolean
    Dim showIconOnly As Boolean
    Dim priority As Variant
    Dim percentileValues As Boolean
    Dim iconSetId As Variant
    '------------------------------------------------------------------------------------------------------
    Dim iconCriteria As VBA.Collection
    Dim dtoCriterium As Scripting.Dictionary
    Dim index As Long
    Dim operator As Long
    Dim value As Variant
    Dim icon As Long
    Dim iconType As Variant
    '------------------------------------------------------------------------------------------------------
    
    stopIfTrue = f.dictionaries.getItem(formatProps, "stopIfTrue")
    showIconOnly = f.dictionaries.getItem(formatProps, "showIconOnly")
    priority = f.dictionaries.getItem(formatProps, "priority")
    percentileValues = f.dictionaries.getItem(formatProps, "percentileValues")
    iconSetId = f.Variables.coalesce(f.dictionaries.getItem(formatProps, "iconSetId"), 7)
    Set iconCriteria = f.dictionaries.getObject(formatProps, "iconCriteria")
    
    Set applyIconSetsFormatCondition = rng.FormatConditions.AddIconSetCondition()
    With applyIconSetsFormatCondition
        .IconSet = iconSetId
        .showIconOnly = showIconOnly
        .priority = priority
        .percentileValues = percentileValues
    End With
    
    For Each dtoCriterium In iconCriteria
        index = f.dictionaries.getItem(dtoCriterium, "index")
        operator = f.dictionaries.getItem(dtoCriterium, "operator")
        value = f.dictionaries.getItem(dtoCriterium, "value")
        icon = f.dictionaries.getItem(dtoCriterium, "icon")
        iconType = f.dictionaries.getItem(dtoCriterium, "type")
        
        With applyIconSetsFormatCondition.iconCriteria(index)
            .operator = operator
            If .value <> value Then .value = value
            .icon = icon
        End With
        
    Next dtoCriterium
    
End Function

Public Function applyExpressionFormatCondition(rng As Excel.Range, formula As String, _
                                                    backColor As Long, Optional fontColor As Long, _
                                                    Optional stopIfTrue As Boolean = False) As Excel.FormatCondition
    Set applyExpressionFormatCondition = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    With applyExpressionFormatCondition
        Call .setFirstPriority
        With .Interior
            .PatternColorIndex = xlAutomatic
            .color = backColor
            .TintAndShade = 0
        End With
        If fontColor Then .Font.color = fontColor
        .stopIfTrue = stopIfTrue
    End With
End Function



Public Sub applyFormatFromDictionary(rng As Excel.Range, properties As Scripting.Dictionary)
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    Dim prop As UIProp
    Dim dictBorders As Scripting.Dictionary
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        Set dict = f.dictionaries.flatDictionary(properties)
        If f.dictionaries.isNonEmptyDictionary(dict) Then
            For Each key In dict.Keys
                Set prop = UIProps.getPropByAlias(VBA.CStr(key))
                If prop Is Nothing Then
                    'skip
                ElseIf VBA.InStr(1, prop.getName, "borders.") Then
                    Stop
                    'borders are set in the separate submethod
                Else
                    Call applyFormatForSingleProperty(rng, prop, prop.getConvertedValue(dict.item(key)))
                End If
            Next key
        End If
        
        '[Borders]
        On Error Resume Next
        Set dictBorders = getBordersDefinitionsFromDict(properties)
        If Not DEV_MODE Then On Error GoTo ErrHandler
        If f.dictionaries.isNonEmptyDictionary(dictBorders) Then Call applyBorders(rng, dictBorders)
        
    End If
    
ErrHandler:
    
End Sub

Private Function getBordersDefinitionsFromDict(dict As Scripting.Dictionary) As Scripting.Dictionary
    Const SEPARATOR_CHARS As String = ",;|"
    '----------------------------------------------------------------------------------------------------------
    Dim keyAlias As Variant
    Dim key As Variant
    Dim subkeys As Variant
    Dim subkey  As Variant
    '----------------------------------------------------------------------------------------------------------
    Dim tempDict As Scripting.Dictionary
    Dim separators As Variant
    Dim varSeparator As Variant
    Dim multiKeyEntry As Boolean
    '----------------------------------------------------------------------------------------------------------
    
    Set getBordersDefinitionsFromDict = f.dictionaries.Create(False)
    
    Set tempDict = f.dictionaries.getObject(dict, UIProps.Borders.getName)
    If f.dictionaries.isNonEmptyDictionary(tempDict) Then
    
        '[Split common border definitions]
        separators = f.strings.toCharactersArray(SEPARATOR_CHARS)
        For Each key In tempDict.Keys
            multiKeyEntry = False
            For Each varSeparator In separators
                If VBA.InStr(1, key, varSeparator) Then
                    multiKeyEntry = True
                    subkeys = VBA.Split(key, varSeparator)
                    If f.arrays.isDefined(subkeys) Then
                        For Each subkey In subkeys
                            Call f.dictionaries.addItem(getBordersDefinitionsFromDict, VBA.trim(subkey), tempDict.item(key))
                        Next subkey
                    End If
                End If
            Next varSeparator

            If Not multiKeyEntry Then
                Call f.dictionaries.addItem(getBordersDefinitionsFromDict, key, tempDict.item(key))
            End If

        Next key
        
    End If
    
End Function

Public Sub applyFormatForSingleProperty(rng As Excel.Range, prop As UIProp, value As Variant)
    
    Select Case prop.getId
        Case UIProps.width.getId:                       rng.EntireColumn.ColumnWidth = value
        Case UIProps.height.getId:                      rng.EntireRow.RowHeight = value
        Case UIProps.minWidth.getId:                    rng.EntireColumn.ColumnWidth = f.Math.getMax(value, rng.EntireColumn.ColumnWidth)
        Case UIProps.minHeight.getId:                   rng.EntireRow.RowHeight = f.Math.getMax(value, rng.EntireRow.RowHeight)
        Case UIProps.BackgroundColor.getId:             rng.Interior.color = value
        Case UIProps.fontSize.getId:                    rng.Font.size = value
        Case UIProps.fontColor.getId:                   rng.Font.color = value
        Case UIProps.FontFamily.getId:                  rng.Font.name = value
        Case UIProps.fontBold.getId:                    rng.Font.Bold = value
        Case UIProps.FontItalic.getId:                  rng.Font.Italic = value
        Case UIProps.VerticalAlignment.getId:           rng.VerticalAlignment = value
        Case UIProps.HorizontalAlignment.getId:         rng.HorizontalAlignment = value
        Case UIProps.IndentLevel.getId:                 rng.IndentLevel = value
        Case UIProps.WrapText.getId:                    rng.WrapText = value
        Case UIProps.ColumnAutoFit.getId:               Call rng.EntireColumn.AutoFit
        Case UIProps.RowAutoFit.getId:                  Call rng.EntireRow.AutoFit
        Case UIProps.NumberFormat.getId:                rng.NumberFormat = value
        Case UIProps.Locked.getId:                      rng.Locked = value
        Case UIProps.PaddingVertical.getId:             Stop: rng.EntireRow.RowHeight = rng.EntireRow.RowHeight + 2 * value
    End Select
    
End Sub





Public Sub applyRowAutoFit(rng As Excel.Range, Optional minHeight As Single)
    Dim heights() As Single
    Dim height As Single
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    Dim subrng As Excel.Range
    Dim dictHeights As Scripting.Dictionary
    Dim colRows As VBA.Collection
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        If rng.Areas.count > 1 Then
            For Each subrng In rng.Areas
                Call applyRowAutoFit(subrng, minHeight)
            Next subrng
        Else
            Set dictHeights = f.dictionaries.Create(False)
            
            With rng
                ReDim heights(1 To .Rows.count)
                For i = LBound(heights) To UBound(heights)
                    heights(i) = .Rows(i).RowHeight
                    Set subrng = f.ranges.join(subrng, .Rows(i))
                Next i
                
                Call subrng.EntireRow.AutoFit
                
                For i = LBound(heights) To UBound(heights)
                    height = f.Math.getMax(heights(i), minHeight)
                    If .Rows(i).RowHeight < height Then
                        Set colRows = f.dictionaries.getObject(dictHeights, height)
                        If colRows Is Nothing Then
                            Set colRows = New VBA.Collection
                            Call dictHeights.Add(height, colRows)
                        End If
                        Call colRows.Add(rng.Rows(i).EntireRow)
                    End If
                Next i
                
                
                
                'Restore minimum height
                For Each key In dictHeights.Keys
                    Set subrng = f.ranges.join(dictHeights.item(key))
                    If f.ranges.IsValid(subrng) Then
                        subrng.EntireRow.RowHeight = VBA.CSng(key)
                    End If
                Next key
                
                'Hide rows that were hidden before
                For i = LBound(heights) To UBound(heights)
                    If heights(i) = 0 Then
                        .Rows(i).EntireRow.Hidden = True
                    End If
                Next i
                
                
            End With
        
        End If
    End If
    
End Sub

Public Sub clearSheetFormatConditions(wks As Excel.Worksheet, Optional protectionPassword As String)
    Dim isProtected As Boolean:             isProtected = wks.ProtectContents
    Dim count As Long
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If isProtected Then Call wks.Unprotect(protectionPassword)
    
    With wks.cells.FormatConditions
        count = .count
        For i = count To 1 Step -1
            Call .item(i).delete
        Next i
    End With
    
    If isProtected Then Call wks.Protect(protectionPassword)
    
End Sub

Public Function getCurrentSelectionAddress(Optional xls As Excel.Application) As String
    If xls Is Nothing Then Set xls = Excel.Application
    On Error Resume Next
    getCurrentSelectionAddress = xls.selection.address
    On Error GoTo 0
End Function

Public Function getNamedRange(Sheet As Excel.Worksheet, name As String) As Excel.Range
    On Error Resume Next
    Set getNamedRange = Sheet.Range(name)
End Function

Public Sub applyColors(dict As Scripting.Dictionary)
    Dim key As Variant
    Dim subcol As VBA.Collection
    Dim rng As Excel.Range
    Dim color As Long
    '------------------------------------------------------------------------------------------------------
    
    For Each key In dict.Keys
        Set subcol = f.dictionaries.getObject(dict, key)
        If f.Collections.isNonEmptyCollection(subcol) Then
            Set rng = f.ranges.join(subcol)
            If f.ranges.IsValid(rng) Then
                color = VBA.CLng(key)
                With rng.Interior
                    If f.colors.isValidRgbColor(color) Then
                        rng.Interior.color = VBA.CLng(key)
                    Else
                        .pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End If
                End With
            End If
        End If
    Next key
    
End Sub



Public Function getVisibleSelection(Optional ByVal xls As Excel.Application) As Excel.Range
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If xls Is Nothing Then Set xls = Excel.Application
    
    On Error Resume Next
    Set rng = xls.selection
    If Not DEV_MODE Then On Error GoTo ErrHandler
    
    If IsValid(rng) Then
        Set getVisibleSelection = rng.SpecialCells(xlCellTypeVisible)
    End If
    
ErrHandler:
    
End Function



Public Function haveCommonParts(rng1 As Excel.Range, rng2 As Excel.Range) As Boolean
    If Not IsValid(rng1) Then
        haveCommonParts = False
    ElseIf Not IsValid(rng2) Then
        haveCommonParts = False
    ElseIf Not rng1.Parent Is rng2.Parent Then
        haveCommonParts = False
    ElseIf rng1.row >= rng2.row + rng2.Rows.count Then
        haveCommonParts = False
    ElseIf rng2.row >= rng1.row + rng1.Rows.count Then
        haveCommonParts = False
    ElseIf rng1.column >= rng2.column + rng2.columns.count Then
        haveCommonParts = False
    ElseIf rng2.column >= rng1.column + rng1.columns.count Then
        haveCommonParts = False
    Else
        haveCommonParts = True
    End If
End Function


Public Function toString(value As Object) As String
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    If value Is Nothing Then
        toString = "Nothing"
    ElseIf Not TypeOf value Is Excel.Range Then
        toString = "Invalid data type"
    Else
        Set rng = value
        If f.ranges.IsValid(rng) Then
            toString = "{Excel.Range} " & rng.address & " (" & rng.Parent.name & ")"
        Else
            toString = "Invalid range"
        End If
    End If
End Function



Public Function findRowByTop(Sheet As Excel.Worksheet, Top As Single) As Long
    Dim i As Long
    Dim rowBottom As Single
    '------------------------------------------------------------------------------------------------------
    
    If Not f.sheets.IsValid(Sheet) Then
        findRowByTop = -1
    ElseIf Top < 0 Then
        findRowByTop = 0
    Else
        For i = 1 To Sheet.Rows.count
            With Sheet.Rows(i)
                rowBottom = .Top + .height
            End With
            If rowBottom >= Top Then
                findRowByTop = i
                Exit For
            End If
        Next i
    End If
    
End Function



Public Function getRangeSizeAndPosition(rng As Excel.Range) As area
    Dim height As Single
    Dim width As Single
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        '[Calculate height]
        If rng.Rows.count = 1 Then
            height = rng.height
        Else
            For i = 1 To rng.Rows.count
                height = height + rng.Rows(i).height
            Next i
        End If
        
        '[Calculate width]
        If rng.columns.count = 1 Then
            width = rng.width
        Else
            For i = 1 To rng.columns.count
                width = width + rng.columns(i).width
            Next i
        End If
        
        With getRangeSizeAndPosition
            .Left = rng.Left
            .Top = rng.Top
            .height = height
            .width = width
        End With
        
    End If
    
End Function




''Author: https://stackoverflow.com/questions/21580795/subtracting-ranges-in-vba-excel
'Public Function subtract(rFirst As Range, rSecond As Range) As Range
'    Dim rInter As Range
'    Dim rReturn As Range
'    Dim rArea As Range
'
'    Set rInter = intersect(rFirst, rSecond)
'    Set mrBuild = Nothing
'
'    If rInter Is Nothing Then 'no overlap
'        Set rReturn = rFirst
'    ElseIf rInter.address = rFirst.address Then 'total overlap
'        Set rReturn = Nothing
'    Else 'partial overlap
'        For Each rArea In rFirst.Areas
'            Set mrBuild = BuildRange(rArea, rInter) 'recursive
'        Next rArea
'        Set rReturn = mrBuild
'    End If
'
'    Set subtract = rReturn
'
'End Function
'
'
'Private Function BuildRange(rArea As Range, rInter As Range, Optional mrBuild As Range = Nothing) As Range
'' Recursive function for SubtractRanges()
''
'' Subtracts rInter from rArea and adds the result to mrBuild
''
'    Dim app As Excel.Application:                       Set app = rArea.Application
'    Dim rLeft As Range, rRight As Range
'    Dim rTop As Range, rBottom As Range
'    Dim rInterSub As Range
'    Dim GoByColumns As Boolean
'
'    Set rInterSub = intersect(rArea, rInter)
'    If rInterSub Is Nothing Then 'no overlap
'        If mrBuild Is Nothing Then
'            Set mrBuild = rArea
'        Else
'            Set mrBuild = app.Union(mrBuild, rArea)
'        End If
'    ElseIf Not rInterSub.address = rArea.address Then 'some overlap
'        If Not rArea.Cells.CountLarge = 1 Then 'just in case there is only one cell for some impossible reason
'
'            ' Decide whether to go by columns or by rows
'            ' (helps when subtracting whole rows/columns)
'            If Not rInterSub.columns.count = rArea.columns.count And _
'            ((Not rInterSub.Cells.CountLarge = 1 And _
'            (rInterSub.rows.count > rInterSub.columns.count _
'            And rArea.columns.count > 1) Or (rInterSub.rows.count = 1 _
'            And Not rArea.columns.count = 1)) Or _
'            (rInterSub.Cells.CountLarge = 1 _
'            And rArea.columns.count > rArea.rows.count)) Then
'                    GoByColumns = True
'            Else
'                    GoByColumns = False
'            End If
'
'            If Not GoByColumns Then
'                Set rTop = rArea.Resize(rArea.rows.count \ 2) 'split the range top to bottom
'                Set rBottom = rArea.Resize(rArea.rows.count - rTop.rows.count).Offset(rTop.rows.count)
'                Set mrBuild = BuildRange(rTop, rInterSub, mrBuild) 'rerun it
'                Set mrBuild = BuildRange(rBottom, rInterSub, mrBuild)
'            Else
'                Set rLeft = rArea.Resize(, rArea.columns.count \ 2) 'split the range left to right
'                Set rRight = rArea.Resize(, rArea.columns.count - rLeft.columns.count).Offset(, rLeft.columns.count)
'                Set mrBuild = BuildRange(rLeft, rInterSub, mrBuild) 'rerun it
'                Set mrBuild = BuildRange(rRight, rInterSub, mrBuild)
'            End If
'        End If
'    End If
'
'    Set BuildRange = mrBuild
'End Function




Public Sub trimSheet(wks As Excel.Worksheet, _
                        Optional lastVisibleRow As Long, Optional hideRows As Boolean = False, _
                        Optional lastVisibleColumn As Long, Optional hideColumns As Boolean = False)
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If lastVisibleRow > 0 And lastVisibleRow < wks.Rows.count Then
        On Error Resume Next
        Set rng = wks.cells(lastVisibleRow, 1).offset(1).Resize(wks.Rows.count - lastVisibleRow).EntireRow
        On Error GoTo 0
        
        If f.ranges.IsValid(rng) Then
            Call rng.clear
            If hideRows Then rng.Hidden = True
        End If
    End If
    
    
    If lastVisibleColumn > 0 And lastVisibleColumn < wks.columns.count Then
        On Error Resume Next
        Set rng = wks.cells(1, lastVisibleColumn).offset(0, 1).Resize(, wks.columns.count - lastVisibleColumn).EntireColumn
        On Error GoTo 0
        
        If f.ranges.IsValid(rng) Then
            Call rng.clear
            If hideColumns Then rng.Hidden = True
        End If
    End If
    
    
End Sub


Public Sub autoFitColumnsWithoutShrinking(rng As Excel.Range)
    Dim widths As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        ReDim widths(1 To rng.columns.count)
        For i = LBound(widths) To UBound(widths)
            widths(i) = rng.columns(i).ColumnWidth
        Next i
    End If
    
    On Error Resume Next
    Call rng.EntireColumn.AutoFit
    On Error GoTo 0
    
    For i = LBound(widths) To UBound(widths)
        If rng.columns(i).ColumnWidth < widths(i) Then
            rng.columns(i).ColumnWidth = widths(i)
        End If
    Next i
    
End Sub

Public Sub autoFitRowsWithoutShrinking(rng As Excel.Range, Optional maxHeight As Single)
    Dim heights As Variant
    Dim i As Long
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        ReDim heights(1 To rng.Rows.count)
        For i = LBound(heights) To UBound(heights)
            heights(i) = rng.Rows(i).RowHeight
        Next i
    End If
    
    On Error Resume Next
    Call rng.EntireRow.AutoFit
    On Error GoTo 0
    
    For i = LBound(heights) To UBound(heights)
        If rng.Rows(i).RowHeight > maxHeight And maxHeight > 0 Then
            rng.Rows(i).RowHeight = maxHeight
        ElseIf rng.Rows(i).RowHeight < heights(i) Then
            rng.Rows(i).RowHeight = heights(i)
        End If
    Next i
    
End Sub


Public Sub expandColumns(rng As Excel.Range, extraWidth As Single)
    Dim i As Long
    Dim width As Single
    '------------------------------------------------------------------------------------------------------
    
    If f.ranges.IsValid(rng) Then
        For i = 1 To rng.columns.count
            With rng.columns(i).EntireColumn
                width = .width + extraWidth
                .width = width
            End With
        Next i
    End If
    
End Sub

Public Function compareRanges(baseRange As Excel.Range, comparedRange As Excel.Range) As Scripting.Dictionary
    Dim rngCommon As Excel.Range
    '------------------------------------------------------------------------------------------------------
    Dim dictBaseAddresses As Scripting.Dictionary
    Dim dictComparedAddresses As Scripting.Dictionary
    Dim addresses As Variant
    Dim key As Variant
    '------------------------------------------------------------------------------------------------------
    Dim colCommon As VBA.Collection
    Dim colOnlyBase As VBA.Collection
    Dim colOnlyCompared As VBA.Collection
    '------------------------------------------------------------------------------------------------------
    
    Set compareRanges = f.dictionaries.Create(False)
    
    If Not f.ranges.IsValid(baseRange) Then
        '[Only compared range is valid]
        If f.ranges.IsValid(comparedRange) Then
            Call compareRanges.Add(ComparisonResult_OnlyInCompared, comparedRange)
        End If
    ElseIf Not f.ranges.IsValid(comparedRange) Then
        '[Only base range is valid]
        If f.ranges.IsValid(baseRange) Then
            Call compareRanges.Add(ComparisonResult_OnlyInBase, baseRange)
        End If
    ElseIf Not baseRange.Parent Is comparedRange.Parent Then
        '[Both ranges valid, but located in different sheets - no common points]
        Call compareRanges.Add(ComparisonResult_OnlyInBase, baseRange)
        Call compareRanges.Add(ComparisonResult_OnlyInCompared, comparedRange)
    Else
        
        Set dictBaseAddresses = createAddressesDictionary(baseRange)
        Set dictComparedAddresses = createAddressesDictionary(comparedRange)
        
        Set colCommon = New VBA.Collection
        Set colOnlyBase = New VBA.Collection
        Set colOnlyCompared = New VBA.Collection
        
        '[Iterate through base addresses dictionary]
        For Each key In dictBaseAddresses.Keys
            If dictComparedAddresses.exists(key) Then
                Call colCommon.Add(dictBaseAddresses.item(key))
                Call dictComparedAddresses.Remove(key)
            Else
                Call colOnlyBase.Add(dictBaseAddresses.item(key))
            End If
        Next key
        
        '[Iterate through compare addresses dictionary]
        For Each key In dictComparedAddresses.Keys
            Call colOnlyCompared.Add(dictComparedAddresses.item(key))
        Next key
        
        With compareRanges
            Call .Add(ComparisonResult_Common, f.ranges.join(colCommon))
            Call .Add(ComparisonResult_OnlyInBase, f.ranges.join(colOnlyBase))
            Call .Add(ComparisonResult_OnlyInCompared, f.ranges.join(colOnlyCompared))
        End With
        
    End If
    
End Function



Private Function createAddressesDictionary(rng As Excel.Range) As Scripting.Dictionary
    Dim wks As Excel.Worksheet
    Dim rngArea As Excel.Range
    Dim row As Long
    Dim col As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim coordinates As String
    '------------------------------------------------------------------------------------------------------
    
    Set createAddressesDictionary = f.dictionaries.Create(False)
    
    Set wks = rng.Parent
    
    For Each rngArea In rng.Areas
        lastRow = f.ranges.getRangeLastRow(rngArea)
        lastCol = f.ranges.getRangeLastColumn(rngArea)
        For row = rngArea.row To lastRow
            For col = rngArea.column To lastCol
                coordinates = row & ";" & col
                Call f.dictionaries.addItem(createAddressesDictionary, coordinates, wks.cells(row, col))
            Next col
        Next row
    Next rngArea
    
    
End Function


Public Sub ungroup(rng As Excel.Range)
    On Error Resume Next
    Call rng.ungroup
End Sub


Public Function getRangeCell(rng As Excel.Range, row As Long, column As Long) As Excel.Range
    If f.ranges.IsValid(rng) Then
        If Not f.Math.isInRange(row, 1, rng.Rows.count) Then
            ' row out of scope
        ElseIf Not f.Math.isInRange(column, 1, rng.columns.count) Then
            ' column out of scope
        Else
            Set getRangeCell = rng.cells(row, column)
        End If
    End If
End Function



Public Sub copyFormulas(rngSource As Excel.Range, rngDestination As Excel.Range)
    Dim cell As Excel.Range
    Dim formula As String
    Dim isArrayFormula As Boolean
    Dim Counter As Long
    '------------------------------------------------------------------------------------------------------
    Dim subrng As Excel.Range
    Dim subCell As Excel.Range
    '------------------------------------------------------------------------------------------------------
    
    If Not f.ranges.IsValid(rngSource) Then
        ' invalid source range
    ElseIf Not f.ranges.IsValid(rngDestination) Then
        ' invalid destination range
    Else
        For Each cell In rngSource.Resize(1).cells
            Counter = Counter + 1
            formula = cell.FormulaR1C1
            isArrayFormula = cell.HasArray
            
            If f.Math.isInRange(Counter, 1, rngDestination.columns.count) Then
                Set subrng = rngDestination.columns(Counter)
                
                If isArrayFormula Then
                    For Each subCell In subrng.cells
                        subCell.FormulaArray = formula
                    Next subCell
                Else
                    subrng.FormulaR1C1 = formula
                End If
                
            End If
            
        Next cell
    
    End If
    
End Sub



Public Sub insertValidationList(rng As Excel.Range, listName As String)
    Const METHOD_NAME As String = "insertValidationList"
    '------------------------------------------------------------------------------------------------------
    Const FORMULA_PATTERN As String = "={0}"
    Dim formula As String
    '------------------------------------------------------------------------------------------------------
    
    If Not DEV_MODE Then On Error GoTo ErrHandler

    formula = f.strings.format(FORMULA_PATTERN, listName)

    If Not f.ranges.IsValid(rng) Then
        Call Exceptions.InvalidRange.raise
    Else
        With rng.validation
            On Error Resume Next
            Call .delete
            If Not DEV_MODE Then On Error GoTo ErrHandler
            
            Call .Add(Type:=xlValidateList, AlertStyle:=xlValidAlertStop, operator:=xlBetween, Formula1:=formula)
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End If

'==========================================================================================================
ExitPoint:
    Dim errNumber As Long, errDescription As String
    If errNumber > 0 Then Call VBA.err.raise(errNumber, CLASS_NAME & "." & METHOD_NAME, errDescription)
    Exit Sub

'----------------------------------------------------------------------------------------------------------
ErrHandler:
    errNumber = VBA.err.number: errDescription = VBA.err.description
    Call ErrorManager.addError(CLASS_NAME, METHOD_NAME, errNumber, errDescription, _
                                "rng; listName", _
                                rng, listName)
    GoTo ExitPoint
    
End Sub

