VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "FUtils"
'----------------------------------------------------------------------------------------------------------
Private Const BACKUP_FOLDER_PATH As String = "E:\dane\VBA\backups"
'----------------------------------------------------------------------------------------------------------

Public Sub printFilePaths()
    Dim addIn As Excel.addIn
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    
    For Each addIn In Application.AddIns2
        If addIn.isOpen Then
            Debug.Print addIn.fullName
        End If
    Next addIn
    
    For Each wkb In Application.Workbooks
        Debug.Print wkb.fullName
    Next wkb
    
End Sub

Public Sub saveAllAddIns( _
                    Optional ByVal skipRepoFiles As Boolean = False, _
                    Optional ByVal saveBackups As Boolean = False, _
                    Optional ByVal saveFolder As String = BACKUP_FOLDER_PATH)
    Const REGEX_MACRO_ENABLED As String = "\.xlsm$"
    '------------------------------------------------------------------------------------------------------
    Dim addIn As Excel.addIn
    Dim wkb As Excel.Workbook
    Dim timestamp As Date:                          timestamp = VBA.Now
    '------------------------------------------------------------------------------------------------------
    
    'Save all add-ins
    For Each addIn In Application.AddIns2
        If Not isAddInSkippedFromSaving(addIn.name, skipRepoFiles) Then
            If addIn.isOpen Then
                On Error Resume Next
                Set wkb = Application.Workbooks(addIn.name)
                On Error GoTo 0
                
                If wkb.name = addIn.name Then
                    Call wkb.save
                    Debug.Print wkb.fullName & " saved"
                    If saveBackups Then
                        Call saveBackup(wkb, timestamp)
                    End If
                End If
            End If
        End If
    Next addIn
    
    'Save all xlam files.
    For Each wkb In Application.Workbooks
        If f.regex.checkIfMatch(wkb.name, REGEX_MACRO_ENABLED) Then
            Call wkb.save
            Debug.Print wkb.fullName & " saved"
            If saveBackups Then
                Call saveBackup(wkb, timestamp)
            End If
        End If
    Next wkb
    
    Debug.Print "Saving completed [" & VBA.format(VBA.Now, "YYYY-MM-DD HH:MM:SS") & "]"
    
End Sub

Private Function isAddInSkippedFromSaving(ByVal addinName As String, ByVal skipRepoFiles As Boolean) As Boolean
    Const REGEX_REPO_FILE As String = "(?:mielk|toolbox)-.*\.xlam$"
    '------------------------------------------------------------------------------------------------------
    If VBA.InStr(1, addinName, "solver", vbTextCompare) Then
        isAddInSkippedFromSaving = True
    ElseIf VBA.InStr(1, addinName, ".xll", vbTextCompare) Then
        isAddInSkippedFromSaving = True
    ElseIf skipRepoFiles Then
        isAddInSkippedFromSaving = f.regex.checkIfMatch(addinName, REGEX_REPO_FILE)
    Else
        isAddInSkippedFromSaving = False
    End If
End Function

Private Sub saveBackup(wkb As Excel.Workbook, timestamp As Date)
    Static fso As Scripting.FileSystemObject
    '------------------------------------------------------------------------------------------------------
    Dim folderName As String
    Dim folderPath As String
    Dim folder As Scripting.folder
    Dim path As String
    '------------------------------------------------------------------------------------------------------
    
    If fso Is Nothing Then
        Set fso = New Scripting.FileSystemObject
    End If
    
    folderName = VBA.format$(timestamp, "YYYY-MM-DD HHMMSS")
    folderPath = f.files.concatenatePath(BACKUP_FOLDER_PATH, folderName)
    If fso.folderExists(folderPath) Then
        Set folder = fso.getFolder(folderPath)
    Else
        Set folder = fso.createFolder(folderPath)
    End If
    path = f.files.concatenatePath(folder.path, wkb.name)
    
    Call fso.CopyFile(wkb.fullName, path, True)
    
End Sub


Public Sub unfreezeExcel()
    With Excel.Application
        .Interactive = True
        .EnableEvents = True
        .ScreenUpdating = True
        .cursor = xlDefault
    End With
End Sub




Public Function convertToAddIn(originalFilePath As String, _
                                Optional ByVal destinationPath As String, _
                                Optional overwriteIfExists As Boolean = False, _
                                Optional ByVal xls As Excel.Application) As String
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    
    If VBA.Len(destinationPath) = 0 Then
        destinationPath = VBA.replace(originalFilePath, ".xlsm", ".xlam")
    End If
    
    If Not overwriteIfExists Then
        destinationPath = f.files.uniqueFilePath(f.files.getUncPath(destinationPath))
    End If
    
    If xls Is Nothing Then Set xls = f.Excel.getBackgroundInstance
    
    Set wkb = f.Books.open_(originalFilePath, False, xls)
    With wkb
        .IsAddin = True
        .CheckCompatibility = False
        Call f.files.deleteFile(destinationPath)
        Call .SaveAs(destinationPath, xlAddIn)
        Call .Close(False)
    End With
    
End Function


Public Sub versionUpdate(version As String, Optional description As String)
    Dim addIn As Excel.addIn
    Dim wkb As Excel.Workbook
    '------------------------------------------------------------------------------------------------------
    
    'Save all add-ins
    For Each addIn In Application.AddIns2
        If Not isAddInSkippedFromSaving(addIn.name, True) Then
            If addIn.isOpen Then
                On Error Resume Next
                Set wkb = Application.Workbooks(addIn.name)
                On Error GoTo 0
                
                If wkb.name = addIn.name Then
                    Call setFileVersion(wkb, version, description)
                    Debug.Print wkb.fullName & " - version number changed to " & version
                End If
            End If
        End If
    Next addIn
    
    Debug.Print "Versioning completed [" & VBA.format(VBA.Now, "YYYY-MM-DD HH:MM:SS") & "]"
    
End Sub

Public Sub displayFileVersions()
    Dim addIn As Excel.addIn
    Dim wkb As Excel.Workbook
    Dim version As String
    Dim message As String
    '------------------------------------------------------------------------------------------------------
    
    'Save all add-ins
    For Each addIn In Application.AddIns2
        If Not isAddInSkippedFromSaving(addIn.name, True) Then
            If addIn.isOpen Then
                On Error Resume Next
                Set wkb = Application.Workbooks(addIn.name)
                On Error GoTo 0
                
                If wkb.name = addIn.name Then
                    version = f.Variables.coalesce(getFileVersion(wkb), "???")
                    message = VBA.String(2, " ") & version & VBA.String(10 - VBA.Len(version), " ") & wkb.fullName
                    Debug.Print message
                End If
            End If
        End If
    Next addIn
    
End Sub
