VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} mergeSelectForm 
   Caption         =   "Merge Files Dialog"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   72
   ClientWidth     =   13308
   OleObjectBlob   =   "mergeSelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "mergeSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public targetWBPath As String
Public targetWSName As String
Public sourceWBPath As String
Public sourceWSName As String

Private Sub cancelButton_Click()
    Unload Me
End Sub

Private Sub okButton_Click()
    If Me.sourceSheetsListBox.Value <> "" & Me.sourceWBPathLabel <> "" & _
        Me.targetSheetsListBox.Value <> "" & Me.targetWBPathLabel <> "" Then
        
        On Error GoTo dump
        Me.targetWBPath = Me.targetWBPathLabel
        Me.targetWSName = Me.targetSheetsListBox.Value
        Me.sourceWBPath = Me.sourceWBPathLabel
        Me.sourceWSName = Me.sourceSheetsListBox.Value
        
        If Me.sourceWBPath <> "" And Me.sourceWSName = "" Then Me.sourceWSName = "RFIs"
        If Me.targetWBPath <> "" And Me.targetWSName = "" Then Me.targetWSName = "RFIs"
        
        On Error GoTo 0
        
        Call MergeTwoFiles
        
        Me.Hide
    Else
        MsgBox "Please select a source workbook and worksheet as well as a target workbook and worksheet.", _
            vbOKOnly & vbCritical, _
            "Incomplete Selection Warning"
    End If
dump:
End Sub


Private Sub MergeTwoFiles()

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim sWB As Workbook
    Dim sWS As Worksheet
    Dim sTable As ListObject
    Set sWB = Application.Workbooks.Open(fileName:=Me.sourceWBPath)
    Set sWS = sWB.Sheets(Me.sourceWSName)
    Set sTable = sWS.ListObjects(1)
    
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim tTable As ListObject
    Set tWB = Application.Workbooks.Open(fileName:=Me.targetWBPath)
    Set tWS = tWB.Sheets(Me.targetWSName)
    Set tTable = tWS.ListObjects(1)
    
    On Error GoTo dump
    Call MergeTables(sTable, tTable)
    On Error GoTo 0
    
    sWB.Close SaveChanges:=False
    tWB.Save
    
dump:
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

End Sub



Private Sub selectSourceWBButton_Click()
    Dim selectedPath As String
    selectedPath = GetAWBPath()
    If selectedPath <> "" Then
        targetWBPath = selectedPath
        Me.sourceWBPathLabel.Caption = selectedPath
        Call PopulateWorksheetList(selectedPath, Me.sourceSheetsListBox)
        
        Me.sourceSheetsListBox.Selected(0) = True
    End If
End Sub

Private Sub selectTargetWBButton_Click()
    Dim selectedPath As String
    selectedPath = GetAWBPath()
    If selectedPath <> "" Then
        targetWBPath = selectedPath
        Me.targetWBPathLabel.Caption = selectedPath
        Call PopulateWorksheetList(selectedPath, Me.targetSheetsListBox)
        
        Me.targetSheetsListBox.Selected(0) = True
    End If
End Sub

Private Sub UserForm_Initialize()

    Me.targetWBPath = ""
    Me.targetWSName = ""
    Me.sourceWBPath = ""
    Me.sourceWSName = ""
    
    Me.versionLabel.Caption = "Merge Reports" & vbCrLf & _
        BidGulp.module_name & " v." & BidGulp.module_version
    Me.Show
End Sub



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


