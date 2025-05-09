Attribute VB_Name = "enumerateCOMAddins"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       10th September 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Sub EntryPointGetCOMAddinReferences()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :
'   Company     :       Alignment Systems Limited
'   Date        :       24th March 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const strMethodName As String = "enumerateCOMAddins.EntryPointGetCOMAddinReferences "
Const cstrConnect As String = "Connected="
Const cstrCreator As String = " Creator="
Const cstrDescription As String = " Description="
Const cstrProgID As String = " ProgID="
Const cstrGUID As String = " GUID="
'Variables
Dim oCOMAddin As OFFICE.COMAddIn
Dim oTargetWorkbook As Excel.Workbook

Set oTargetWorkbook = ThisWorkbook

For Each oCOMAddin In oTargetWorkbook.Application.COMAddIns
    Debug.Print cstrConnect & oCOMAddin.Connect & cstrCreator & oCOMAddin.Creator & _
    cstrDescription & "[" & oCOMAddin.Description & "]" & cstrGUID & oCOMAddin.GUID & cstrProgID & oCOMAddin.ProgID
Next

End Sub


