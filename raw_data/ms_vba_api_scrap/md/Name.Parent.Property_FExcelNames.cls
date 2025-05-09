VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FExcelNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FExcelNames"
    '----------------------------------------------------------------------------------------------------------

Public Function addValue(Parent As Variant, name As String, ByVal value As Variant) As Excel.name
    If Len(value) = 0 Then value = "_"
    Set addValue = Parent.names.Add(name, value)
End Function

Public Function getValue(Parent As Object, name As String) As Variant
    Dim objName As Excel.name
    '------------------------------------------------------------------------------------------------------
    Set objName = get_(Parent, name)
    If Not objName Is Nothing Then
        If checkIfNameHasConstValue(objName) Then
            getValue = evaluateExcelNameConstValue(objName)
        Else
            getValue = evaluateExcelNameRangeValue(objName)
        End If
    End If
End Function

Public Function get_(Parent As Object, name As String) As Excel.name
    Dim rangeName As String
    '------------------------------------------------------------------------------------------------------
    rangeName = getFullPath(Parent, name)
    If exists(Parent, name) Then
        Set get_ = Parent.names(rangeName)
    End If
End Function

Public Function getFullPath(Parent As Object, name As String) As String
    Const EXCEL_NAME_PATTERN As String = "{0}!{1}"
    '------------------------------------------------------------------------------------------------------
    Dim parentName As String
    '------------------------------------------------------------------------------------------------------
    If TypeOf Parent Is Excel.Workbook Then
        getFullPath = name
    ElseIf TypeOf Parent Is Excel.Worksheet Then
        parentName = VBA.IIf(VBA.InStr(1, Parent.name, " "), "'" & VBA.replace(Parent.name, "'", "''") & "'", Parent.name)
        getFullPath = f.strings.format(EXCEL_NAME_PATTERN, parentName, name)
    End If
End Function

Public Function exists(Parent As Object, name As String) As Boolean
    Dim rangeName As String
    Dim objName As Excel.name
    Dim names As Variant
    '------------------------------------------------------------------------------------------------------
    
    rangeName = getFullPath(Parent, name)
    
    If VBA.Len(rangeName) Then
        On Error Resume Next
        Set names = Parent.names
        On Error GoTo 0
        
        If Not VBA.IsEmpty(names) Then
            For Each objName In Parent.names
                If f.strings.compareStrings(objName.name, rangeName) Then
                    exists = True
                    Exit For
                End If
            Next objName
        End If
    End If
    
End Function

Private Function checkIfNameHasConstValue(excelName As Excel.name) As Boolean
    checkIfNameHasConstValue = f.strings.compareStrings(excelName.RefersTo, excelName.RefersToR1C1)
End Function

Private Function evaluateExcelNameConstValue(namedRange As Excel.name) As Variant
    Const REGEX_PATTERN As String = "^(?:=\""*|)(.*?)(?:\""|)$"
    '------------------------------------------------------------------------------------------------------
    Dim strValue As String
    '------------------------------------------------------------------------------------------------------
    strValue = f.regex.getFirstGroupMatch(VBA.CStr(namedRange.value), REGEX_PATTERN)
    If VBA.isNumeric(strValue) Then
        evaluateExcelNameConstValue = strValue * 1
    ElseIf f.Variables.isBoolean(strValue, allowStringRepresentation:=True) Then
        evaluateExcelNameConstValue = VBA.CBool(strValue)
    Else
        evaluateExcelNameConstValue = strValue
    End If
End Function

Private Function evaluateExcelNameRangeValue(namedRange As Excel.name) As Variant
    Dim rng As Excel.Range
    '------------------------------------------------------------------------------------------------------
    Set rng = f.ranges.getFirstCell(getReferredRange(namedRange))
    evaluateExcelNameRangeValue = rng.value
End Function

Public Function getReferredRange(excelName As Excel.name) As Excel.Range
    Dim wks As Excel.Worksheet
    Dim sheetName As String
    Dim address As String
    '------------------------------------------------------------------------------------------------------
    
    With excelName
        sheetName = VBA.replace(f.strings.substring(.RefersTo, "=", "!"), "'", vbNullString)
        address = f.strings.substring(.RefersTo, "!", "")
        
        On Error Resume Next
        If TypeOf excelName.Parent Is Excel.Worksheet Then
            Set wks = excelName.Parent
        ElseIf TypeOf excelName.Parent Is Excel.Workbook Then
            Set wks = excelName.Parent.Worksheets(sheetName)
        End If
        On Error GoTo 0
        
        If Not wks Is Nothing Then
            Set getReferredRange = wks.Range(address)
        End If
        
    End With

End Function


Public Function assignRange(Parent As Object, name As String, rng As Excel.Range, _
                 Optional visible As Boolean = True) As Excel.name
    Dim fullName As String:             fullName = getFullPath(Parent, name)
    Dim excelName As Excel.name
    '------------------------------------------------------------------------------------------------------

    Set assignRange = get_(Parent, name)

    If f.ranges.IsValid(rng) Then
        If assignRange Is Nothing Then
            rng.name = fullName
        Else
            With assignRange
                .visible = visible
                .RefersTo = rng
            End With
        End If
    ElseIf Not assignRange Is Nothing Then
        Call delete(Parent, name)
    End If
        
End Function

Public Sub assignValue(Parent As Object, name As String, value As Variant, _
                                                                     Optional IsVisible As Boolean = False)
    Dim fullName As String:             fullName = getFullPath(Parent, name)
    Dim excelName As Excel.name
    '------------------------------------------------------------------------------------------------------

    Set excelName = get_(Parent, name)
    If Not excelName Is Nothing Then
        Call delete(Parent, name)
    End If
        
    Call Parent.names.Add(name:=name, RefersTo:=value, visible:=IsVisible)

End Sub



Public Function getWorkbookExcelNames(wkb As Excel.Workbook) As Variant
    Dim en As Excel.name
    Dim i As Long
    Dim arr As Variant
    '------------------------------------------------------------------------------------------------------
    
    If f.Books.IsValid(wkb) Then
        If wkb.names.count Then
            ReDim arr(1 To 5, 1 To wkb.names.count)
            For Each en In wkb.names
                i = i + 1
                arr(1, i) = en.name
                arr(2, i) = en.RefersTo
                arr(3, i) = en.value
                arr(4, i) = en.Comment
                If Not en.Parent Is Nothing Then arr(5, i) = en.Parent.name
            Next en
        End If
    End If
    
    getWorkbookExcelNames = arr
    
End Function



'[DELETING]
Public Sub delete(Parent As Object, name As String)
    Dim Events As Boolean
    '------------------------------------------------------------------------------------------------------

    With Parent.Application
        Events = .EnableEvents
        .EnableEvents = False
    End With

    If exists(Parent, name) Then
        Call Parent.names(getFullPath(Parent, name)).delete
    End If

    Parent.Application.EnableEvents = Events

End Sub

Public Sub deleteAll(ParamArray parents() As Variant)
    Dim varParent As Variant
    Dim name As Excel.name
    '------------------------------- -----------------------------------------------------------------------
    For Each varParent In parents
        If TypeOf varParent Is Excel.Worksheet Or TypeOf varParent Is Excel.Workbook Then
            For Each name In varParent.names
                On Error Resume Next
                If name.Parent Is varParent Then
                    Call name.delete
                End If
            Next name
        End If
    Next varParent
End Sub

