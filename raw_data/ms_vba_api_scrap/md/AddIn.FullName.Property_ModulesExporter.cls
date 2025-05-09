VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModulesExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "ModulesExporter"
'----------------------------------------------------------------------------------------------------------
Private Const BUILD_IN_ADDINS_FOLDER As String = "C:\Program Files\Microsoft Office\"
Private Const DESTINATION_FOLDER As String = "D:\Dropbox\tm\mielk\mielk-vba\github"
'----------------------------------------------------------------------------------------------------------


Public Sub exportModulesFromOpenFiles()
    Dim wkb As Excel.Workbook
    Dim addIn As Excel.addIn
    '------------------------------------------------------------------------------------------------------
    
    For Each wkb In Excel.Workbooks
        Call F.Developer.exportModules(wkb, DESTINATION_FOLDER)
    Next wkb
    
    For Each addIn In Excel.Application.AddIns2
        If VBA.InStr(1, addIn.FullName, BUILD_IN_ADDINS_FOLDER) = 0 Then
            Set wkb = Excel.Workbooks(addIn.name)
            Call F.Developer.exportModules(wkb, DESTINATION_FOLDER)
            Debug.Print addIn.FullName
        End If
    Next addIn

End Sub
