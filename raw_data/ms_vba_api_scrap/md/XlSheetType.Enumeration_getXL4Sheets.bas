Attribute VB_Name = "getXL4Sheets"
Option Explicit
Option Compare Text
Option Base 0
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Private Function DecodeSheetVisible(Test As Excel.XlSheetVisibility) As String
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Select Case Test
    Case xlSheetHidden
        DecodeSheetVisible = "Hidden"
    Case xlSheetVeryHidden
        DecodeSheetVisible = "VeryHidden"
    Case xlSheetVisible
        DecodeSheetVisible = "Visible"
End Select

End Function

Private Function DecodeSheetType(Test As Excel.XlSheetType) As String
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================

Select Case Test
    Case xlExcel4IntlMacroSheet
        DecodeSheetType = "Excel4Intl"
    Case xlExcel4MacroSheet
        DecodeSheetType = "Excel4Macro"
End Select

End Function


Function Tester()
'============================================================================================================================
'
'
'   Author      :       John Greenan
'   Email       :       john.greenan@alignment-systems.com
'   Company     :       Alignment Systems Limited
'   Date        :       6th April 2015
'
'   Purpose     :       Matching Engine in Excel VBA for Alignment Systems Limited
'
'   References  :       See VB Module FL for list extracted from VBE
'   References  :
'============================================================================================================================
Const strSheetName As String = "SheetName="
Const strVisible As String = "Visibility="

Dim oSheet As Excel.Worksheet
Dim oWorkbook As Excel.Workbook

For Each oSheet In Application.Excel4IntlMacroSheets
    With oSheet
        If .Type = Excel.XlSheetType.xlExcel4IntlMacroSheet Then
            Debug.Print strSheetName & .Name, DecodeSheetType(.Type), strVisible & DecodeSheetVisible(.Visible)
        End If
    End With
Next

For Each oSheet In Application.Excel4MacroSheets
    With oSheet
        If .Type = Excel.XlSheetType.xlExcel4MacroSheet Then
            Debug.Print strSheetName & .Name, DecodeSheetType(.Type), strVisible & DecodeSheetVisible(.Visible)
        End If
    End With
Next


End Function



