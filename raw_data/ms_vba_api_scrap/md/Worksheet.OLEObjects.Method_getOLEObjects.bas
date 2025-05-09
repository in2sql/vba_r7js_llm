Attribute VB_Name = "getOLEObjects"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       5th April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Function EntryPointListOLEObjects()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       5th April 2014
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
'Constants
Const strMethodName As String = "getOLEObjects.EntryPointListOLEObjects "
Const strWorkSheetName As String = "WorksheetName="
Const strTypeName As String = " TypeName="
Const strProgID As String = " ProgID="
Const strVisible As String = " Visible?="
Const strObjectName As String = " Name="
Const strObjectObjectCaption As String = " Caption="
'Variables
Dim oWorkSheet As Excel.Worksheet
Dim oWorkbook As Excel.Workbook
Dim oObject As Excel.OLEObject
Dim strPrintString As String

Set oWorkbook = ThisWorkbook

For Each oWorkSheet In oWorkbook.Worksheets
    If oWorkSheet.OLEObjects.count = 0 Then
        Debug.Print oWorkSheet.Name & "--> has no OLEObjects"
    Else
        For Each oObject In oWorkSheet.OLEObjects
            strPrintString = strWorkSheetName & oWorkSheet.Name & strTypeName & TypeName(oObject) & strProgID & oObject.ProgID & strVisible & oObject.Visible & strObjectObjectCaption & oObject.Object.Caption
            
            If oObject.Name <> "" Then
                strPrintString = strPrintString & strObjectName & oObject.Name
            End If
            
            Debug.Print strPrintString
            strPrintString = ""
        Next
    End If
Next

End Function

