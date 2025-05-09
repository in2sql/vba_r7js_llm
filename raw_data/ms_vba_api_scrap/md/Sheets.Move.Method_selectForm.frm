VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} selectForm 
   Caption         =   "Select Form"
   ClientHeight    =   4140
   ClientLeft      =   216
   ClientTop       =   780
   ClientWidth     =   6636
   OleObjectBlob   =   "selectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "selectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' SelectForm is a dialog used to select an HTML and Workbook/Worksheet for adding additional
' Bidder RFI comments (to an existing report).

Public htmlPath As String
Public targetWBPath As String
Public targetWSName As String


Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub okButton_Click()
    If Me.worksheetListBox.Value <> "" & Me.selectedWBPathLabel <> "" & Me.selectedHTMLPathLabel <> "" Then
        Me.targetWSName = Replace(Me.worksheetListBox.Value, "$", "")
        Me.targetWBPath = Me.selectedWBPathLabel.Caption
        Me.htmlPath = Me.selectedHTMLPathLabel.Caption
        Me.Hide
    Else
        MsgBox "Make sure you've selected an HTML report, target Workbook and Worksheet first.", _
            vbOKOnly & vbCritical, _
            "Incomplete Selection Warning"
    End If
End Sub

Private Sub selectHTMLButton_Click()
    Dim selectedPath As String
    selectedPath = GetAnHTMLPath()
    If selectedPath <> "" Then
        htmlPath = selectedPath
        Me.selectedHTMLPathLabel.Caption = selectedPath
    End If
End Sub

Private Sub selectWorkbookButton_Click()
    Dim selectedPath As String
    selectedPath = GetAWBPath()
    If selectedPath <> "" Then
        targetWBPath = selectedPath
        Me.selectedWBPathLabel.Caption = selectedPath
        Call PopulateWorksheetList(selectedPath, Me.worksheetListBox)
        
        Me.worksheetListBox.Selected(0) = True
    End If
End Sub

Private Function GetAnHTMLPath() As String
    ' Returns EMPTY if user cancels, otherwise returns path string
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "HTML", "*.htm?"
        .Title = "Choose an HTML file"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetAnHTMLPath = .SelectedItems(1)
    End With
End Function

Private Function GetAWBPath() As String
    Dim fd As FileDialog, wbPath As String
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls?"
        .Title = "Choose an Excel file"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        GetAWBPath = .SelectedItems(1)
    End With
End Function

Public Sub PopulateWorksheetList(wbPath As String, aListBox As MSForms.ListBox)

    If wbPath <> "" Then
        ' REFERENCE: Microsoft ActiveX Data Objects 6.1 Library
        Dim cn As ADODB.Connection
        Dim rsSheets As ADODB.Recordset
        Dim sName As String
        Set cn = New ADODB.Connection
    
        ' https://www.connectionstrings.com/
        cn.ConnectionString = _
            "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & wbPath & ";" & _
            "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';"
        cn.Open
        
        ' Get a list of all the Worksheets as a Recordset
        Set rsSheets = cn.OpenSchema(adSchemaTables)
        Do Until rsSheets.EOF
            sName = Replace(rsSheets.Fields("TABLE_NAME").Value, "$", "")
            If LCase(sName) <> "instructions" Then
                aListBox.AddItem sName
            End If
            rsSheets.MoveNext
        Loop
        rsSheets.Close
        cn.Close
    End If
        
End Sub

Private Sub UserForm_Initialize()
    Me.versionLabel.Caption = "Add to Existing Report" & vbCrLf & _
        BidGulp.module_name & " v." & BidGulp.module_version
    Me.Show
End Sub

