'     m
'    / \    part of
'   |...|
'   |...|   bottle light quality management software
'   |___|   by error on line 1 (erroronline.one) available on https://github.com/erroronline1/qualitymanagement
'   / | \
'set selectedLanguage in the Workbook_Open-sub and customize your language and location specific values in the Locale module

'be aware that changing the tables layout will have an effect on the code as well.
'make sure to update sheets and Locale values on either side

Option Explicit
Public selectedLanguage As String
Public parentPath

Private Sub Workbook_Open()
    selectedLanguage = "D"

    'get parent path to vb_libraries to be imported
    Dim path() As String
    path() = Split(ThisWorkbook.path, "\")
    Dim i As Integer
    parentPath = ""
    For i = 0 To UBound(path) - 2 'according to upward steps in folder hierarchy
        parentPath = parentPath & path(i) & "\"
    Next i
    'load essentials as module and execute opening procedure
    Dim Essentials
    Set Essentials = CreateObject("Scripting.Dictionary")
    Essentials.Add "Essentials", parentPath & "vb_library\Admin_Essentials.vba"

    If importModules(Essentials) Then asyncOpen
End Sub

Private Sub asyncOpen()
    Essentials.openRoutine
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Essentials.monitorRowsColumns Sh, Nothing
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
    Essentials.monitorRowsColumns Sh, target
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Essentials.closeRoutine SaveAsUI, Cancel
End Sub

Public Function importModules(ByVal libraries As Object) As Boolean
    Dim lib As Variant, modloop As Variant
    On Error Resume Next
    'Application.DisplayAlerts = False
    'rename existing modules and import external modules
    For Each lib In libraries
        If Len(Dir(libraries(lib))) > 0 Then
            With ThisWorkbook.VBProject.VBComponents
                'renaming to _old because sometimes modules are removed on finishing of the code only, resulting in enumeration of module names
                .Item(lib).Name = lib & "_OLD"
                .Import libraries(lib)
            End With
        ElseIf Len(Dir(libraries(lib))) = 0 And lib = "Specific" and Len(Dir(ThisWorkbook.parentPath & "vb_library\" & "Admin_RecordDocument.vba")) Then
            'if no specific module containing ThisWorkbook.Name as specified in Essentials by default it is presumably a record document template
            With ThisWorkbook.VBProject.VBComponents
                'renaming to _old because sometimes modules are removed on finishing of the code only, resulting in enumeration of module names
                .Item(lib).Name = lib & "_OLD"
                .Import ThisWorkbook.parentPath & "vb_library\" & "Admin_RecordDocument.vba"
            End With
        End If
    Next lib
    'if no external modules have been found rename existing modules to default name
    For Each lib In libraries
        Dim loaded As Boolean
        For Each modloop In ThisWorkbook.VBProject.VBComponents
            If (modloop.Name = lib) Then loaded = True: Exit For
        Next modloop
        If loaded Then
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(lib & "_OLD")
        Else
            ThisWorkbook.VBProject.VBComponents(lib & "_OLD").Name = lib
        End If
    Next lib
    Application.DisplayAlerts = True
    On Error GoTo 0
    importModules = True
End Function
