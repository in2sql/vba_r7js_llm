Attribute VB_Name = "BidGulp"
'***********************************************************************
'                            Module Metadata
'***********************************************************************
Public Const module_name As String = "BidGulp"
Public Const module_author As String = "Ben Fisher"
Public Const module_version As String = "1.0.7"
Public Const module_date As Date = #5/31/2024#
Public Const module_notes As String = _
    "As is typical, in a rush to get usable code on the street, so " & _
    "opportunities for refactoring abound. But for now, it'll do."
Public Const module_license As String = "Created by Ben Fisher, GNU General Public License, v3.0"


'***********************************************************************
'                          Referenced Libraries
'***********************************************************************
' 1. Microsoft XML, 6.0
' 2. Microsoft HTML Object Library
' 3. Microsoft Scripting Runtime
' 4. Microsoft VBScript Regular Expresions 5.5
' 5. Microsoft Visual Basic for Applications Extensibility 5.3

'***********************************************************************
'                            User Preferences
'***********************************************************************
' Table Placement on Worksheet
Public Const METADATA_TARGET_CELL As String = "A1"
Public Const TABLE_TARGET_CELL As String = "A10"

' Typical Table Prefer
Public Const VERTICALPADDING = 10
Public Const TYPICALFONTSIZE = 9

' Table Color Preferences (references webcolors)
Public Const METAINTERIORCOLOR = webcolors.LEMONCHIFFON
Public Const HEADERINTERIORCOLOR = webcolors.PALETURQUOISE
Public Const TABLERULECOLOR = webcolors.STEELBLUE

Public Const OFFICEOFCOUNSELCOLOR = webcolors.LIMEGREEN
Public Const CONTRACTINGCOLOR = webcolors.PURPLE
Public Const LESSONSLEARNEDCOLOR = webcolors.ORANGERED

Public Const TIMESTAMPFILE = False

'***********************************************************************
'                       Table Configuration Enums
'***********************************************************************
' Enum types must appear before procedures.

Public Const TABLENAME As String = "BidderRFIs"

Public Const RIGHTHOLLOWARROW As Long = 9655
Public Const LEFTHOLLOWARROW As Long = 9665

Public Enum FieldNos
    ItemNo = 1
    CommentID = 2
    CommentDiscipline = 3
    CommentSheet = 4
    CommentDetail = 5
    CommentSpec = 6
    commentText = 7
    CommentClassification = 0   'Current Not Used
    AssignedDiscipline = 8
    AssignedParty = 9
    ResponseDiscussion = 10
    PreliminaryResponse = 11
    RequiresAmendment = 12
    EngineeringFinalResponse = 13
    HasTechLeadQA = 14
    TechServicesFinalResponse = 15
    ContractingResponse = 16
    EngineeringResponseToContracting = 17
    OfficeOfCounselResponse = 18
    EngineeringResponseToOfficeOfCounsel = 19
    LessonsLearnedCapture = 20
    [_First] = ItemNo
    [_Last] = LessonsLearnedCapture
End Enum

Public Enum Widths
    XXS = 5
    XS = 10
    SM = 12
    MD = 20
    LG = 25
    XL = 30
    XXL = 45
    XXXL = 72
End Enum

'***********************************************************************
'                      Paths, Workbooks, and Sheets
'***********************************************************************

Public Function GetHTMLPath() As String
    ' Returns EMPTY if user cancels, otherwise returns path string
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "HTML", "*.htm?"
        .Title = "Choose an HTML file"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetHTMLPath = .SelectedItems(1)
    End With
End Function

Public Function GetFolderPath() As String
    ' Returns EMPTY if user cancels, otherwise returns path string
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetFolderPath = .SelectedItems(1)
    End With
End Function

Function CreateWorkbook(save_path As String, _
    Optional workbook_name As String = "Bidder RFI Summary Report", _
    Optional include_timestamp As Boolean = TIMESTAMPFILE) As Workbook
    ' Return a  new Workbook object with the provided name
    ' and appends with a timestamp as noted.
    Dim combined_workbook As Workbook
    Dim file_name As String
    Set combined_workbook = Workbooks.Add
    Application.DisplayAlerts = False
    With combined_workbook
        .Title = workbook_name
        If include_timestamp Then
            file_name = save_path & "\" & workbook_name & " " _
                & Format(Now(), "YYYY-MM-DD hh-mm-ss") & ".xlsx"
        Else
            file_name = save_path & "\" & workbook_name & ".xlsx"
        End If
        .SaveAs fileName:=file_name, FileFormat:=xlOpenXMLWorkbook
    End With
    Application.DisplayAlerts = True
    Set CreateWorkbook = combined_workbook
End Function

Function IterateSheetName(baseName As String)
    Dim maxIndex As Long
    For Each sht In ActiveWorkbook.Sheets
        If Left(sht.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
    Next
    If maxIndex = 0 Then
        IterateSheetName = baseName
    Else
        IterateSheetName = baseName & maxIndex
    End If
End Function

Function RenameSheet(ByVal target_sheet As Worksheet, ByVal sheet_name As String) As String
    ' Renames a Worksheet (Spreadsheet tab)
    ' It limits the name to the maximum permitted character count (31 - 4 = 27) and removes
    ' illegal characters from the name
    
    Dim illegal_characters As Variant
    Dim new_sheet_name As String
    Dim i As Long
    'Create array of characters that are not permitted in worksheet names
    illegal_characters = Array("/", "\", "?", "*", ":", "[", "]")
    If Len(sheet_name) > 27 Then
        new_sheet_name = Left(sheet_name, 27)
    Else
        new_sheet_name = sheet_name
    End If
    For i = LBound(illegal_characters) To UBound(illegal_characters)
        new_sheet_name = Replace(new_sheet_name, illegal_characters(i), "")
    Next
    On Error GoTo dump
    target_sheet.Name = IterateSheetName(new_sheet_name)
    RenameSheet = target_sheet.Name
dump:
End Function


'***********************************************************************
'                            Parse to Array
'***********************************************************************
Function ParseToArray(namedConstant As String) As Variant
    ParseToArray = Split(namedConstant, ", ")
End Function

Function ParseToLongArray(namedConstant As String) As Variant
    Dim arr As Variant
    Dim arr2 As Variant
    Dim i As Long
    
    arr = ParseToArray(namedConstant)
    ReDim arr2(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        arr2(i) = CLng(arr(i))
    Next
    ParseToLongArray = arr2
End Function


'***********************************************************************
'                          Regular Expressions
'***********************************************************************
Public Function MatchesPattern(phrase As String, search_string As String) As Boolean
    Dim regex As New RegExp
    Dim match_collection As MatchCollection
    Dim a_match As Match
    With regex
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
    End With
    regex.Pattern = phrase
    Set match_collection = regex.Execute(search_string)
    If match_collection.Count <> 0 Then MatchesPattern = True
End Function

'***********************************************************************
'                          Global Table Styles
'***********************************************************************
Function StyleExists(styleName As String) As Boolean
    Dim dummyBool As Boolean
    On Error GoTo dump
        If ActiveWorkbook.TableStyles(styleName) = styleName Then dummyBool = True
        ' Simply need to evaluate an expression to force potential error
    On Error GoTo 0
    StyleExists = True
    Exit Function
dump:
End Function

Public Function CreateBasicTableStyle() As String
    Dim i As Long
    Dim styleName As String: styleName = "Simple Style"
    
    If Not StyleExists(styleName) Then
        ActiveWorkbook.TableStyles.Add styleName
        ActiveWorkbook.TableStyles(styleName).ShowAsAvailableTableStyle = True
    End If
    
    With ActiveWorkbook.TableStyles(styleName)
        With .TableStyleElements(xlHeaderRow)
            'Interior must preceed Font colors for Table Styles
            .Interior.Color = HEADERINTERIORCOLOR
            .Font.Color = ContrastText(.Interior.Color)
            .Font.FontStyle = "Bold"
            'NOTE: You cannot edit the header row height or font size in the
            ' table style definition. For this reason, those parameters are
            ' edited in the method called ApplyTableStyle().
            With .Borders(xlInsideVertical)
                .Color = webcolors.WHITE
                .Weight = xlThin
            End With
        End With
        .TableStyleElements(xlRowStripe1).Clear
        With .TableStyleElements(xlRowStripe1)
            For i = xlEdgeTop To xlEdgeTop
                With .Borders(i)
                    .Color = TABLERULECOLOR
                    .Weight = xlThin
                End With
            Next
            With .Borders(xlInsideVertical)
                .Color = RGB(230, 230, 230)
                .Weight = xlThin
            End With
        End With
        .TableStyleElements(xlRowStripe2).Clear
        With .TableStyleElements(xlRowStripe2)
            For i = xlEdgeTop To xlEdgeTop
                With .Borders(i)
                    .Color = TABLERULECOLOR
                    .Weight = xlThin
                End With
            Next
            With .Borders(xlInsideVertical)
                .Color = RGB(230, 230, 230)
                .Weight = xlThin
            End With
        End With
    End With
    CreateBasicTableStyle = styleName
End Function

Public Sub ApplySimpleStyleToTable(aTable As ListObject)

    aTable.TableStyle = ""
    aTable.TableStyle = CreateBasicTableStyle()
    
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        With aTable.ListRows(1).Range
            .Clear
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignTop
        End With
        temporaryRow = True
    Else
        With aTable.DataBodyRange.Rows
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignTop
        End With
    End If
    aTable.DataBodyRange.WrapText = True
    
    Application.ScreenUpdating = False
    Dim centeredColumns As Variant
    centeredColumns = Array(FieldNos.AssignedDiscipline, FieldNos.AssignedParty, _
                            FieldNos.RequiresAmendment, FieldNos.HasTechLeadQA)
    
    For i = LBound(centeredColumns) To UBound(centeredColumns)
        aTable.ListColumns(centeredColumns(i)).DataBodyRange.HorizontalAlignment = xlHAlignCenter
    Next
    If temporaryRow Then aTable.ListRows.Item(1).Delete
    
    Application.ScreenUpdating = True
    
End Sub

Function IntersectsTable(proposedRange As Range) As Boolean
    Dim sht As Worksheet
    Dim lstObj As ListObject
    Dim doesIntersect As Boolean    'False by default
    For Each lstObj In ActiveSheet.ListObjects
        If Not Intersect(proposedRange, lstObj.Range) Is Nothing Then doesIntersect = True
    Next
    IntersectsTable = doesIntersect
End Function

Function RangeToTable(headerRange As Range, sht As Worksheet, _
    Optional tblName As String = TABLENAME) As ListObject

    Dim newName As String
    newName = AutoincrementTableName(tblName)

    If IntersectsTable(headerRange) = False Then
        sht.ListObjects.Add(xlSrcRange, headerRange, , xlYes).Name = newName
        Set RangeToTable = sht.ListObjects(newName)
    End If
End Function

Function TableHasData(aTable As ListObject) As Boolean
    If Not aTable.DataBodyRange Is Nothing Then TableHasData = True
End Function

Function AutoincrementTableName(baseName As String)
    Dim maxIndex As Long
    Dim sht As Worksheet
    Dim lstObj As ListObject
    For Each sht In ActiveWorkbook.Sheets
        For Each lstObj In sht.ListObjects
            If Left(lstObj.Name, Len(baseName)) = baseName Then maxIndex = maxIndex + 1
        Next
    Next
    If maxIndex = 0 Then AutoincrementTableName = baseName Else AutoincrementTableName = baseName & maxIndex
End Function

'***********************************************************************
'                    Metadata & Table Initialization
'***********************************************************************
Public Sub WriteMetaData(sht As Worksheet, _
    Optional projName As String = "Project: <Type in Project Name>")
    
    Application.ScreenUpdating = False
    With sht.Range(METADATA_TARGET_CELL)
        .Value = "RFI WORKSHEET"
        .Font.Name = "Arial Black"
        .Font.Size = 18
        .EntireRow.AutoFit
    End With
    With sht.Range(METADATA_TARGET_CELL).Offset(1, 0)
        .Value = projName
        .Font.Name = "Arial Black"
        .Font.Size = 14
        .EntireRow.AutoFit
    End With
    With sht.Range(METADATA_TARGET_CELL).Offset(2, 0)
        .Value = "Report created " & Format(Now, "dd MMM YYYY at hh:mm")
        .VerticalAlignment = xlVAlignTop
        .EntireRow.RowHeight = 20
    End With
   
   
    Dim FullWidthRange As Range
    Set FullWidthRange = sht.Range(METADATA_TARGET_CELL, _
                     sht.Range(METADATA_TARGET_CELL).Offset(8, FieldNos.[_Last] - 1))
    With FullWidthRange
        .HorizontalAlignment = xlCenterAcrossSelection
        .Interior.Color = METAINTERIORCOLOR
        .Font.Color = ContrastText(.Interior.Color)
    End With
    
    Dim arr As Variant
    arr = Array("PM:", "TL:", "TS:", "COR:", "CS:")
    With sht.Range(METADATA_TARGET_CELL).Offset(3, 0).Resize(UBound(arr) + 1, 1)
        .Value = WorksheetFunction.Transpose(arr)
        .HorizontalAlignment = xlHAlignLeft
        .Font.Name = "Arial Black"
    End With
   
    Set arr = Nothing
    arr = Array("<PM Name>", "<TL Name>", "<TS POC Name>", "<COR Name>", "<CS Name>")
    With sht.Range(METADATA_TARGET_CELL).Offset(3, 1).Resize(UBound(arr) + 1, 1)
        .Value = WorksheetFunction.Transpose(arr)
        .HorizontalAlignment = xlHAlignLeft
    End With
    
    With sht.Range(METADATA_TARGET_CELL).Offset(8, 1)
        .Value = ChrW(RIGHTHOLLOWARROW)
        .HorizontalAlignment = xlHAlignRight
        .Font.Color = webcolors.SADDLEBROWN
    End With
    
    With sht.Range(METADATA_TARGET_CELL).Offset(8, 6)
        .Value = ChrW(LEFTHOLLOWARROW) & " Expand for Detailed Data"
        .HorizontalAlignment = xlHAlignLeft
        .Font.Color = webcolors.SADDLEBROWN
    End With
    
    With sht.Range(METADATA_TARGET_CELL).Offset(8, FieldNos.[_Last] - 1)
        .Value = "Created using " & module_name & " v." & module_version
        .HorizontalAlignment = xlHAlignRight
        .Font.Size = 8
        .Font.Color = webcolors.SADDLEBROWN
    End With
    
    Set FullWidthRange = Nothing
    Application.ScreenUpdating = True
End Sub

Public Sub WriteTableHeaders(sht As Worksheet)
    Application.ScreenUpdating = False
    Dim ColumnNames As Variant
    ColumnNames = Array( _
                        "Item No", _
                        "Comment ID", _
                        "Discipline", _
                        "Sheet", _
                        "Detail", _
                        "Spec", _
                        "Comment Text", _
                        "Assigned Discipline", _
                        "Assigned Party", _
                        "Response Discussions", _
                        "Preliminary Response", _
                        "Amend Required", _
                        "Engineering Final Response", _
                        "TL QA Reviewed", _
                        "Tech Services Final Response", _
                        "Contracting Response", _
                        "Engineering Response to Contracting", _
                        "OC Response", _
                        "Engineering Response to OC", _
                        "JED LL Capture")
    With sht.Range(TABLE_TARGET_CELL).Resize(1, UBound(ColumnNames) + 1)
        .Value = ColumnNames
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignCenter
    End With
    Set ColumnNames = Nothing
    Application.ScreenUpdating = True
End Sub

Public Sub FormatHeaderRow(sht As Worksheet)
    Application.ScreenUpdating = False
    Dim HeaderRow As Range
    Set HeaderRow = sht.Range(sht.Range(TABLE_TARGET_CELL), _
                                      sht.Range(TABLE_TARGET_CELL).Offset(0, FieldNos.[_Last] - 1))
    Dim HeaderRowNumber As Long
    HeaderRowNumber = HeaderRow.Row
    HeaderRow.ColumnWidth = Widths.XL
    HeaderRow(1, FieldNos.commentText).ColumnWidth = Widths.XXXL
    Dim NarrowerColumns As Range
    Set NarrowerColumns = Union(sht.Range(Cells(HeaderRowNumber, FieldNos.CommentID), _
                                Cells(HeaderRowNumber, FieldNos.CommentSpec)), _
                                Cells(HeaderRowNumber, FieldNos.AssignedDiscipline), _
                                Cells(HeaderRowNumber, FieldNos.AssignedParty), _
                                Cells(HeaderRowNumber, FieldNos.RequiresAmendment), _
                                Cells(HeaderRowNumber, FieldNos.HasTechLeadQA))
    NarrowerColumns.ColumnWidth = Widths.XS
    Set NarrowerColumns = Nothing
    Cells(HeaderRowNumber, FieldNos.ItemNo).ColumnWidth = Widths.XXS
    Cells(HeaderRowNumber, FieldNos.RequiresAmendment).ColumnWidth = Widths.XS
    Cells(HeaderRowNumber, FieldNos.HasTechLeadQA).ColumnWidth = Widths.XS
    With HeaderRow
        .Interior.Color = HEADERINTERIORCOLOR
        .Font.Color = ContrastText(.Interior.Color)
        With .Borders(xlEdgeBottom)
            .Color = TABLERULECOLOR
            .Weight = xlMedium
        End With
    End With
    With Cells(HeaderRowNumber, FieldNos.ContractingResponse)
        .Interior.Color = CONTRACTINGCOLOR
        .Font.Color = ContrastText(.Interior.Color)
    End With
    With Cells(HeaderRowNumber, FieldNos.OfficeOfCounselResponse)
        .Interior.Color = OFFICEOFCOUNSELCOLOR
        .Font.Color = ContrastText(.Interior.Color)
    End With
    With Cells(HeaderRowNumber, FieldNos.LessonsLearnedCapture)
        .Interior.Color = LESSONSLEARNEDCOLOR
        .Font.Color = ContrastText(.Interior.Color)
    End With
    
    With HeaderRow
        .WrapText = False
        .EntireRow.AutoFit
        .WrapText = True
        .RowHeight = .RowHeight + VERTICALPADDING
        .VerticalAlignment = xlVAlignCenter
        .HorizontalAlignment = xlHAlignCenter
    End With
    
    Set HeaderRow = Nothing
    Application.ScreenUpdating = True
End Sub

Public Function InitializeTable(sht As Worksheet) As ListObject
    Dim HeaderRow As Range
    Set HeaderRow = sht.Range(sht.Range(TABLE_TARGET_CELL), _
                                      sht.Range(TABLE_TARGET_CELL).Offset(0, FieldNos.[_Last] - 1))
    Set InitializeTable = RangeToTable(headerRange:=HeaderRow, sht:=sht, tblName:=TABLENAME)
End Function



'***********************************************************************
'                Basis Validation and Conditional Format
'***********************************************************************
Sub InsertDropdown(aTable As ListObject, _
                    targetColumn As Variant, _
                    selectionSet As String, _
                    Optional suppressError As Boolean = False)
    'NOTE: This method inserts a validation list with values parsed from the selectionSet, into
    ' the cells of the targetColumn as dropdown lists. USE: combine with conditional formatting.
    ' selectionSet must be a single string with options seperated by a comma and space.
            
    ' Test for empty table
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        aTable.ListRows(1).Range.Interior.Color = xlNone
        aTable.ListRows(1).Range.Font.Bold = False
        aTable.ListRows(1).Range.Font.Color = webcolors.BLACK
        temporaryRow = True
    End If
    
    With aTable.ListColumns(targetColumn).DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=selectionSet
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = "Disallowed Input"
        .InputMessage = ""
        .ErrorMessage = "Please select from the options: " & selectionSet
        .ShowInput = True
        .ShowError = suppressError
    End With
    
    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete
End Sub


'***********************************************************************
'                       Advanced Table Formatting
'***********************************************************************
Public Sub ApplyDefaultValues(aTable As ListObject)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If Not aTable.DataBodyRange Is Nothing Then
        For i = 1 To aTable.DataBodyRange.Rows.Count
            If aTable.DataBodyRange(i, FieldNos.RequiresAmendment).Value = "" Then aTable.DataBodyRange(i, FieldNos.RequiresAmendment).Value = "TBD"
            If aTable.DataBodyRange(i, FieldNos.HasTechLeadQA).Value = "" Then aTable.DataBodyRange(i, FieldNos.HasTechLeadQA).Value = "No"
        Next i
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Sub ApplyStatusFormats(aTable As ListObject, _
                        targetColumn As Variant, _
                        selectionSet As String, _
                        formatSets As Collection)
    'NOTE: This method inserts conditional formatting based on the validation "dropdown" list
    ' values in the targetColumn. Values must match those in selectionSet (not case sensitive).
    ' hasSecondaryFormats highlights the whole row with accent scheme, while the main
    ' condition only highlights the targetColumn values. The user must manually update
    ' preferences here in this method.
    
    ' Test for empty table
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        With aTable.ListRows(1).Range
            .Font.Bold = False
            .Font.Color = webcolors.BLACK
            .Font.Size = TYPICALFONTSIZE
        End With
        temporaryRow = True
    End If
    
    Dim choices As Variant
    Dim i As Long

    choices = ParseToArray(selectionSet)
    For i = LBound(choices) To UBound(choices)
        choices(i) = """" & choices(i) & """"
    Next

    Dim statusColumn As Range
    Set statusColumn = aTable.ListColumns(targetColumn).DataBodyRange

    Dim firstCell As String
    firstCell = "$" & Replace(statusColumn(1).Address, "$", "")

    statusColumn.FormatConditions.Delete

    For i = LBound(choices) To UBound(choices)
        statusColumn.FormatConditions.Add Type:=xlExpression, _
            Formula1:="=IF(LOWER(" & firstCell & ")=" & choices(i) & ",TRUE,FALSE)"
    Next

    For i = 1 To formatSets.Count
        With statusColumn.FormatConditions(i)
            .Interior.Color = formatSets(i)("interior")
            .Font.Color = formatSets(i)("font")
            .Font.Bold = formatSets(i)("bold")
        End With
    Next

    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete

End Sub

Sub MuteDefaultDiscipline(aTable As ListObject, _
                          targetColumn As Variant)
    'NOTE: This Proc is based on the principles of ApplyStatusFormats(), except that it
    ' specifically looks for cells with a space as the first character, and then applies
    ' conditional formatting that makes the font a shade of gray.
    
    ' Test for empty table
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        With aTable.ListRows(1).Range
            .Font.Bold = False
            .Font.Color = webcolors.BLACK
            .Font.Size = TYPICALFONTSIZE
        End With
        temporaryRow = True
    End If
    
    Dim statusColumn As Range
    Set statusColumn = aTable.ListColumns(targetColumn).DataBodyRange

    Dim firstCell As String
    firstCell = "$" & Replace(statusColumn(1).Address, "$", "")

    statusColumn.FormatConditions.Delete

    ' NOTE: ChrW(32) is a space character
    statusColumn.FormatConditions.Add Type:=xlExpression, _
            Formula1:="=IF(LEFT(" & firstCell & ", 1)="" "",TRUE,FALSE)"

    With statusColumn.FormatConditions(1)
        '<GRAY>
        .Interior.Color = RGB(235, 235, 235)
        .Font.Color = webcolors.DARKSLATEGRAY
        .Font.Bold = False
    End With

    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete

End Sub

Sub HighlightContractingCell(aTable As ListObject, _
                          targetColumn As Variant, _
                          taskerColumn As Variant, _
                          formatSets As Collection)
    'NOTE: This Proc is based on the principles of ApplyStatusFormats(), except that it
    ' specifically looks at a non-empty cell to trigger the condition, and then applies
    ' conditional formatting as input.
    
    ' Test for empty table
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        With aTable.ListRows(1).Range
            .Font.Bold = False
            .Font.Color = webcolors.BLACK
            .Font.Size = TYPICALFONTSIZE
        End With
        temporaryRow = True
    End If
    
    Dim statusColumn As Range
    Set statusColumn = aTable.ListColumns(targetColumn).DataBodyRange

    Dim checkColumn As Range
    Set checkColumn = aTable.ListColumns(taskerColumn).DataBodyRange

    Dim firstCell As String
    firstCell = "$" & Replace(checkColumn(1).Address, "$", "")
    
    Dim statusFirstCell As String
    statusFirstCell = "$" & Replace(statusColumn(1).Address, "$", "")


    statusColumn.FormatConditions.Delete
    statusColumn.FormatConditions.Add Type:=xlExpression, _
        Formula1:="=IF(OR(TRIM(" & firstCell & ")=""Contracting"", " & statusFirstCell & "<>""""),TRUE,FALSE)"
            
    
    For i = 1 To formatSets.Count
        With statusColumn.FormatConditions(i)
            .Interior.Color = formatSets(i)("interior")
            .Font.Color = formatSets(i)("font")
            .Font.Bold = formatSets(i)("bold")
        End With
    Next

    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete

End Sub

Sub HighlightTaskerCell(aTable As ListObject, _
                          targetColumn As Variant, _
                          taskerColumn As Variant, _
                          formatSets As Collection)
    'NOTE: This Proc is based on the principles of ApplyStatusFormats(), except that it
    ' specifically looks at a non-empty cell to trigger the condition, and then applies
    ' conditional formatting as input.
    
    ' Test for empty table
    Dim temporaryRow As Boolean
    If TableHasData(aTable) = False Then
        aTable.ListRows.Add
        With aTable.ListRows(1).Range
            .Font.Bold = False
            .Font.Color = webcolors.BLACK
            .Font.Size = TYPICALFONTSIZE
        End With
        temporaryRow = True
    End If
    
    Dim statusColumn As Range
    Set statusColumn = aTable.ListColumns(targetColumn).DataBodyRange

    Dim checkColumn As Range
    Set checkColumn = aTable.ListColumns(taskerColumn).DataBodyRange

    Dim firstCell As String
    firstCell = "$" & Replace(checkColumn(1).Address, "$", "")

    statusColumn.FormatConditions.Delete
    statusColumn.FormatConditions.Add Type:=xlExpression, _
            Formula1:="=IF(" & firstCell & "<>"""",TRUE,FALSE)"
    
    For i = 1 To formatSets.Count
        With statusColumn.FormatConditions(i)
            .Interior.Color = formatSets(i)("interior")
            .Font.Color = formatSets(i)("font")
            .Font.Bold = formatSets(i)("bold")
        End With
    Next

    ' Remove dummy row needed for adding to empty table
    If temporaryRow Then aTable.ListRows.Item(1).Delete

End Sub


Public Sub AddAllFormattedDropdowns(aTable As ListObject)
    
    InsertDropdown aTable:=aTable, targetColumn:=FieldNos.AssignedDiscipline, _
        selectionSet:="Installation, Civil, Geotech, Environmental, Architecture, Structural, Mechanical, Plumbing, Electrical, Comm, Cyber, Contracting, Specs", _
        suppressError:=True
        
    MuteDefaultDiscipline aTable:=aTable, targetColumn:=FieldNos.AssignedDiscipline
    
    InsertDropdown aTable:=aTable, targetColumn:=FieldNos.AssignedParty, selectionSet:="AE(DBB), Contracting, TL, PM, SME, MCX, COS, PDT(DB)"
    
    'Create colors schemes
    Dim alert As New Dictionary
    Dim warn As New Dictionary
    Dim good As New Dictionary
    Dim muted As New Dictionary
    
    alert.Add Key:="interior", Item:=webcolors.ORANGERED
    alert.Add Key:="font", Item:=ContrastText(alert("interior"))
    alert.Add Key:="bold", Item:=True

    warn.Add Key:="interior", Item:=webcolors.YELLOW
    warn.Add Key:="font", Item:=ContrastText(warn("interior"))
    warn.Add Key:="bold", Item:=False

    good.Add Key:="interior", Item:=webcolors.LIMEGREEN    'RGB(0, 176, 80)
    good.Add Key:="font", Item:=ContrastText(good("interior"))
    good.Add Key:="bold", Item:=True

    muted.Add Key:="interior", Item:=RGB(235, 235, 235)
    muted.Add Key:="font", Item:=webcolors.DARKSLATEGRAY
    muted.Add Key:="bold", Item:=False
    
    Dim formats As Collection
    
    'Create dropdowns and conditional formats for Requires Amendment column
    Dim choices As String
    Set formats = New Collection
    formats.Add alert
    formats.Add good
    formats.Add muted
        
    choices = "Yes, No, TBD"
    InsertDropdown aTable:=aTable, targetColumn:=FieldNos.RequiresAmendment, selectionSet:=choices
    ApplyStatusFormats aTable, targetColumn:=FieldNos.RequiresAmendment, selectionSet:=choices, formatSets:=formats
    
    'Create dropdowns and conditional formats for Has Tech Lead QA column
    Set formats = New Collection
    formats.Add alert
    formats.Add good
    
    choices = "No, Yes"
    InsertDropdown aTable:=aTable, targetColumn:=FieldNos.HasTechLeadQA, selectionSet:=choices
    ApplyStatusFormats aTable, targetColumn:=FieldNos.HasTechLeadQA, selectionSet:=choices, formatSets:=formats
    
    Set alert = Nothing
    Set warn = Nothing
    Set good = Nothing
    Set muted = Nothing
    
    'Create formatting for Contracting response
    Set formats = New Collection
    
    Dim tasker As New Dictionary
    tasker.Add Key:="interior", Item:=webcolors.LAVENDER
    tasker.Add Key:="font", Item:=ContrastText(tasker("interior"))
    tasker.Add Key:="bold", Item:=False
    
    formats.Add tasker
    HighlightContractingCell aTable:=aTable, targetColumn:=FieldNos.ContractingResponse, _
        taskerColumn:=FieldNos.AssignedDiscipline, formatSets:=formats
    
    'Create formatting for OC response
    tasker.RemoveAll
    tasker.Add Key:="interior", Item:=webcolors.HONEYDEW
    tasker.Add Key:="font", Item:=ContrastText(tasker("interior"))
    tasker.Add Key:="bold", Item:=False
    
    Set formats = New Collection
    formats.Add tasker
    HighlightTaskerCell aTable:=aTable, targetColumn:=FieldNos.OfficeOfCounselResponse, _
        taskerColumn:=FieldNos.OfficeOfCounselResponse, formatSets:=formats
    
    'Create formatting for JED LL Capture response
    tasker.RemoveAll
    tasker.Add Key:="interior", Item:=webcolors.MISTYROSE
    tasker.Add Key:="font", Item:=ContrastText(tasker("interior"))
    tasker.Add Key:="bold", Item:=False
    
    Set formats = New Collection
    formats.Add tasker
    HighlightTaskerCell aTable:=aTable, targetColumn:=FieldNos.LessonsLearnedCapture, _
        taskerColumn:=FieldNos.LessonsLearnedCapture, formatSets:=formats

End Sub

'***********************************************************************
'                      Report Creation Procedures
'***********************************************************************

Public Sub ReformatSheet(sht As Worksheet)
    
    Application.ScreenUpdating = False
    ActiveWindow.DisplayGridlines = False
    
    sht.Cells.Rows.AutoFit
    sht.Cells.Font.Size = TYPICALFONTSIZE
    
    WriteMetaData sht, projName:="Project: <Type In Project Name Here>"
    WriteTableHeaders sht
    
    Dim aTable As ListObject
    Set aTable = sht.ListObjects(1)
    
    FormatHeaderRow sht
    
    ApplySimpleStyleToTable aTable
    AddAllFormattedDropdowns aTable

    Application.ScreenUpdating = True
End Sub

Public Sub DestructiveReformatSheet(sht As Worksheet)

    Application.ScreenUpdating = False
    ActiveWindow.DisplayGridlines = False
    
    On Error Resume Next
    sht.Cells.Ungroup
    On Error GoTo 0
    
    sht.Cells.Clear
    sht.Cells.Rows.AutoFit
    sht.Cells.Font.Size = TYPICALFONTSIZE
    
    WriteMetaData sht, projName:="Project: <Type In Project Name Here>"
    WriteTableHeaders sht
    
    Dim aTable As ListObject
    Set aTable = InitializeTable(sht)
    FormatHeaderRow sht
    
    ApplySimpleStyleToTable aTable
    AddAllFormattedDropdowns aTable
    Application.ScreenUpdating = True
End Sub

Public Sub ComfyRows(aTable As ListObject)
    
    aTable.DataBodyRange.Rows.WrapText = False
    aTable.DataBodyRange.Rows.AutoFit
    aTable.DataBodyRange.Rows.WrapText = True
    For Each aRow In aTable.DataBodyRange.Rows
        aRow.RowHeight = aRow.RowHeight + VERTICALPADDING
    Next

End Sub

Public Sub ParseReportHTML(fPath As String, sht As Worksheet)
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False

'    Dim sTime As Long, eTime As Long
'    sTime = Timer

    Dim http As MSXML2.XMLHTTP60
    Dim html As MSHTML.HTMLDocument
        
    Set http = New XMLHTTP60
    Set html = New HTMLDocument
    
    On Error GoTo invalid_error:
    With http
        .Open "GET", fPath, False
        .Send
    End With
    
    html.body.innerHTML = http.responseText
    Set http = Nothing
    
    Dim projName As String
    projName = html.getElementsByClassName("reportSubHeader").Item(0).innerText
    projName = Trim(Left(projName, InStr(1, projName, "Review") - 1))
    sht.Range("A2").Value = projName

    Dim cnt As Long, i As Long
    cnt = html.getElementsByTagName("td").Length
    If cnt = 0 Then
        Set http = Nothing
        Set html = Nothing
        MsgBox "File has no comments.", vbInformation & vbOK, "No Comments in File Dialog"
        Exit Sub
    End If
    
    ReDim tdIDs(cnt - 1) As String
    For i = 0 To cnt - 1
        tdIDs(i) = html.getElementsByTagName("td").Item(i).innerText
    Next
    
    Dim dict As New Dictionary
    For i = 0 To cnt - 1
        If MatchesPattern("^\d{7,8}$", tdIDs(i)) Then
            dict.Add i, i
        End If
    Next
    
    Dim commentsCount As Long
    commentsCount = dict.Count
    
    ReDim cmts(commentsCount - 1) As String
    For i = 0 To commentsCount - 1
        cmts(i) = html.getElementsByClassName("report_comment").Item(i).innerText
    Next
    
    Set html = Nothing
    
    Dim tagIDs() As Long
    ReDim tagIDs(0 To commentsCount - 1)
    
    Dim aTable As ListObject
    Set aTable = sht.ListObjects(1)
           
    Dim aRow As ListRow
    For i = 0 To commentsCount - 1
        tagIDs(i) = dict.Items(i)
        Set aRow = aTable.ListRows.Add
        With aRow
            aRow.Range(1, FieldNos.ItemNo).Value = i + 1                                ' Number
            aRow.Range(1, FieldNos.CommentID).Value = tdIDs(tagIDs(i))                  ' DrChecks ID
            aRow.Range(1, FieldNos.CommentDiscipline).Value = tdIDs(tagIDs(i) + 1)      ' Discipline
            aRow.Range(1, FieldNos.CommentSheet).Value = tdIDs(tagIDs(i) + 2)           ' Sheet
            aRow.Range(1, FieldNos.CommentDetail).Value = tdIDs(tagIDs(i) + 3)          ' Detail
            aRow.Range(1, FieldNos.CommentSpec).Value = tdIDs(tagIDs(i) + 4)            ' Spec
            aRow.Range(1, FieldNos.commentText).Value = cmts(i)                         ' Comment Text
        
            aRow.Range(1, FieldNos.commentText).Offset(0, 1).Value = _
                " " & tdIDs(tagIDs(i) + 1)
        End With
    Next
    
    Call ComfyRows(aTable)
    Call ApplyDefaultValues(aTable)
    
    sht.Range("C1:F1").EntireColumn.Group
    sht.Outline.ShowLevels 1, 1
    
'    eTime = Timer
'
'    MsgBox "Total time: " & Format((eTime - sTime), "0.00") & _
'        " seconds to do " & commentsCount & " comments." & vbCrLf & _
'        "Time per comment: " & Format((eTime - sTime) / commentsCount, "0.000") & _
'        " sec/comment.", vbInformation & vbOkay, "Success"

    Set html = Nothing
    Set http = Nothing
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
invalid_error:
    Debug.Print "Invalid File or File Path."
    Set html = Nothing
    Set http = Nothing
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub


Private Sub OverwriteReportToActiveSheet()
    DestructiveReformatSheet ActiveSheet
    ParseReportHTML GetHTMLPath, ActiveSheet
End Sub

Private Sub CheckWBs()

    For Each wb In Application.Workbooks
        Debug.Print wb.Name
    Next
    Debug.Print ThisWorkbook.Name
End Sub


Public Sub WriteToNewFile()

    Dim htmlPath As String
    htmlPath = GetHTMLPath()

    If htmlPath <> "" Then
        Dim wb As Workbook
        Dim sourceWB As Workbook
        
        Dim sht As Worksheet
        Dim instrSht As Worksheet
    
        Dim fso As FileSystemObject
        Set fso = New FileSystemObject
        
        Dim fileName As String
        Dim friendlyName As String: friendlyName = "Bidder RFI Summary"
        Dim timeStamp As String
            
        Set wb = Workbooks.Add
        Set sourceWB = ThisWorkbook 'Should be BidGulp
        Set instrSht = sourceWB.Sheets("Instructions")
        
        Application.DisplayAlerts = False
        
        With wb
            .Title = friendlyName
            If TIMESTAMPFILE Then
                fileName = fso.GetParentFolderName(htmlPath) & "\" & _
                            friendlyName & " " & _
                            Format(Now(), "YYYY-MM-DD hh-mm-ss") & ".xlsx"
            Else
                fileName = fso.GetParentFolderName(htmlPath) & "\" & _
                            friendlyName & ".xlsx"
            End If
            
            
            
            DestructiveReformatSheet wb.Sheets(1)
            ParseReportHTML htmlPath, wb.Sheets(1)
            
            wb.Sheets(1).Name = IterateSheetName("RFIs")
            instrSht.Copy Before:=wb.Sheets(1)
            wb.Sheets(1).Visible = True
            
            wb.Sheets("RFIs").Activate
            
            Dim pocD As POCDialog
            Set pocD = New POCDialog
            
            If fso.FileExists(fileName) Then
                fileName = fso.GetParentFolderName(htmlPath) & "\" & _
                            friendlyName & " " & _
                            Format(Now(), "YYYY-MM-DD hh-mm-ss") & ".xlsx"
            End If
            Set fso = Nothing
            
            .SaveAs fileName:=fileName, FileFormat:=xlOpenXMLWorkbook
            Set pocD = Nothing
        End With
        
        Set wb = Nothing
        Set sht = Nothing
        
    End If
End Sub



Public Sub AddToExistingFile()

    Dim selectDialog As selectForm
    Set selectDialog = New selectForm

    Dim htmlPath As String
    Dim wbPath As String, wb As Workbook
    Dim wsName As String
    Dim aTable As ListObject
    
    htmlPath = selectDialog.htmlPath
    wbPath = selectDialog.targetWBPath
    wsName = selectDialog.targetWSName
    
    Set selectDialog = Nothing
    
    If htmlPath <> "" And wbPath <> "" And wsName <> "" Then
    
        Set wb = Application.Workbooks.Open(fileName:=wbPath)
        Set ws = wb.Sheets(wsName)
        Set aTable = ws.ListObjects(1)
        
            ' Create a Dictionary listing all the existing comment IDs
        Dim eComments As Dictionary
        Set eComments = New Dictionary
        If Not aTable.DataBodyRange Is Nothing Then
            For Each aComment In aTable.ListColumns("Comment ID").DataBodyRange
                If Not eComments.Exists(aComment) Then eComments.Add CStr(aComment), CStr(aComment)
            Next
        End If
        
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    
        ' Load HTML file and comments
        Dim http As MSXML2.XMLHTTP60
        Dim html As MSHTML.HTMLDocument
    
        Set http = New XMLHTTP60
        Set html = New HTMLDocument
    
        On Error GoTo invalid_error:
        With http
            .Open "GET", htmlPath, False
            .Send
        End With
    
        html.body.innerHTML = http.responseText
        Set http = Nothing
    
        Dim cnt As Long, i As Long
        cnt = html.getElementsByTagName("td").Length
        If cnt = 0 Then
            Set http = Nothing
            Set html = Nothing
            MsgBox "File has no comments.", vbInformation & vbOK, "No Comments in File Dialog"
            Exit Sub
        End If
    
        ReDim tdIDs(cnt - 1) As String
        For i = 0 To cnt - 1
            tdIDs(i) = html.getElementsByTagName("td").Item(i).innerText
        Next
    
        Dim dict As New Dictionary
        For i = 0 To cnt - 1
            If MatchesPattern("^\d{7,8}$", tdIDs(i)) Then
                dict.Add i, i
            End If
        Next
    
        Dim commentsCount As Long
        commentsCount = dict.Count
    
        ReDim cmts(commentsCount - 1) As String
        For i = 0 To commentsCount - 1
            cmts(i) = html.getElementsByClassName("report_comment").Item(i).innerText
        Next
    
        Set html = Nothing
    
        Dim tagIDs() As Long
        ReDim tagIDs(0 To commentsCount - 1)
    
        Dim aRow As ListRow, j As Long
        For i = 0 To commentsCount - 1
            tagIDs(i) = dict.Items(i)
            If Not eComments.Exists(tdIDs(tagIDs(i))) Then
                Set aRow = aTable.ListRows.Add
                With aRow
                    aRow.Range(1, FieldNos.ItemNo).Value = eComments.Count + j + 1              ' Number
                    aRow.Range(1, FieldNos.CommentID).Value = tdIDs(tagIDs(i))                  ' DrChecks ID
                    aRow.Range(1, FieldNos.CommentDiscipline).Value = tdIDs(tagIDs(i) + 1)      ' Discipline
                    aRow.Range(1, FieldNos.CommentSheet).Value = tdIDs(tagIDs(i) + 2)           ' Sheet
                    aRow.Range(1, FieldNos.CommentDetail).Value = tdIDs(tagIDs(i) + 3)          ' Detail
                    aRow.Range(1, FieldNos.CommentSpec).Value = tdIDs(tagIDs(i) + 4)            ' Spec
                    aRow.Range(1, FieldNos.commentText).Value = cmts(i)                         ' Comment Text
    
                    aRow.Range(1, FieldNos.commentText).Offset(0, 1).Value = _
                        " " & tdIDs(tagIDs(i) + 1)
    
                    j = j + 1
                End With
            End If
        Next
    
        Call ComfyRows(aTable)
        Call ApplyDefaultValues(aTable)
        
        Dim pocD As POCDialog
        Set pocD = New POCDialog
        
        wb.Save
        Set pocD = Nothing
    End If
    
    Set html = Nothing
    Set http = Nothing

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Dispose objects
    Set eComments = Nothing
    Set aTable = Nothing
    Set sht = Nothing
    Exit Sub
invalid_error:
    Debug.Print "Invalid File or File Path."
    Set html = Nothing
    Set http = Nothing

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Dispose objects
    Set eComments = Nothing
    Set aTable = Nothing
    Set sht = Nothing

End Sub


Public Sub MergeTables(sourceTable As ListObject, targetTable As ListObject, _
    Optional mergeFields As String = "Response Discussions, Preliminary Response")
   
    Dim srcIDRange As Range
    Dim srcIDDict As Dictionary
    Set srcIDRange = sourceTable.ListColumns("Comment ID").DataBodyRange
    Set srcIDDict = New Dictionary
    For Each anID In srcIDRange
        srcIDDict.Add Key:=anID.Value, Item:=anID.Row - Range(TABLE_TARGET_CELL).Row
    Next
    
    Dim tarIDRange As Range
    Dim tarIDDict As Dictionary
    Set tarIDRange = targetTable.ListColumns("Comment ID").DataBodyRange
    Set tarIDDict = New Dictionary
    For Each anID In tarIDRange
        tarIDDict.Add Key:=anID.Value, Item:=anID.Row - Range(TABLE_TARGET_CELL).Row
    Next
    
    Dim sColumn As Range, sRowIndex As Long
    Dim tColumn As Range, tRowIndex As Long
    Dim allMergeFields As Variant
    
    allMergeFields = ParseToArray(mergeFields)
    On Error Resume Next
    For Each aField In allMergeFields
        Set sColumn = sourceTable.ListColumns(aField).DataBodyRange
        Set tColumn = targetTable.ListColumns(aField).DataBodyRange
        For Each anID In srcIDDict.Keys
            If tarIDDict.Exists(anID) Then
                sRowIndex = srcIDDict(anID)
                rRowIndex = tarIDDict(anID)
                If sColumn(sRowIndex).Value <> "" And tColumn(rRowIndex).Value = "" Then
                    With tColumn(rRowIndex)
                        .Value = sColumn(sRowIndex).Value
                        .Font.Color = webcolors.ORANGERED
                    End With
                End If
            End If
        Next
    Next
    On Error GoTo 0
    
    Set srcIDDict = Nothing
    Set tarIDDict = Nothing
End Sub

Public Sub MergeFiles()

    Dim mergeForm As mergeSelectForm
    Set mergeForm = New mergeSelectForm
    
    Unload mergeForm
    
End Sub



'***********************************************************************
'                          Developer Procedures
'***********************************************************************
Private Sub UpdateVersionNumber()
    For Each sht In ThisWorkbook.Sheets
        If sht.Name = "Macros" Then
            With ThisWorkbook.Sheets("Macros")
                .Unprotect Password:=""
                With .Range("K3")
                    .Value = module_name & " v." & module_version
                    .HorizontalAlignment = xlHAlignRight
                    .Font.Size = 9
                    .Font.Italic = True
                End With
                .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                        AllowFiltering:=True, UserInterfaceOnly:=True, Password:=""
                .EnableSelection = xlUnlockedCells
                .EnableOutlining = True
            End With
        Else
            With sht
                .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                        AllowFiltering:=True, UserInterfaceOnly:=True, Password:=""
                .EnableSelection = xlUnlockedCells
                .EnableOutlining = True
            End With
        End If
    Next



End Sub

Private Sub UnprotectSheets()
    For Each sht In ThisWorkbook.Sheets
        sht.Unprotect Password:=""
    Next
End Sub

