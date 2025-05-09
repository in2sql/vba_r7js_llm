Attribute VB_Name = "Blad_Functies"
'===============================================
'@details       Scheet functions for excel
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================
Sub LockUpSheet()
    If ActiveSheet.ProtectContents = True Then ActiveSheet.Unprotect
    ActiveSheet.UsedRange.Locked = True
    ActiveSheet.UsedRange.FormulaHidden = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    ActiveSheet.EnableSelection = xlNoSelection
End Sub

Sub SetGrid1CM()
    Cells.ColumnWidth = 4.29
    
    Cells.RowHeight = 28.25

End Sub

Sub SetGrid2CM()
    Cells.ColumnWidth = 4.29 * 2
    
    Cells.RowHeight = 28.25 * 2

End Sub

Sub QuickBorder()
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row + Selection.Rows.Count - 1, Selection.Column))
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
    End With
    With Range(Cells(Selection.Row, Selection.Column), Cells(Selection.Row, Selection.Column + Selection.Columns.Count - 1))
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
    End With
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThick
    End With
    
End Sub

Sub CalculateCell()
    Selection.Calculate
End Sub

Sub MergeFormulaToOneCell()

    Dim strPattern As String: strPattern = "[a-zA-Z]{1,2}[0-9]{1,2}"
    Dim strReplace As String: strReplace = ""
    Dim regEx As New RegExp
    Dim strInput As String
    Dim Myrange As Range
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    
    While regEx.test(Selection.Formula)
        adjust = False
        Set array1 = regEx.Execute(Selection.Formula)
        For Each a In array1
            If InStr(ActiveSheet.Range(a.Value).Formula, "=") Then
                Selection.Formula = Replace(Selection.Formula, a.Value, "(" & Replace(ActiveSheet.Range(a.Value).Formula, "=", "") & ")")
                adjust = True
            End If
        Next a
        If Not adjust Then
            Exit Sub
        End If
    
    Wend

End Sub

Sub MultiColToSingleCol()

    startrow = Selection.Row
    activerow = startrow
    startcol = Selection.Column
    endrow = Selection.Row + Selection.Rows.Count - 1
    endcol = Selection.Column + Selection.Columns.Count - 1
    
    Ic = 0
    Ir = 0
    
    While Ic <> Selection.Columns.Count
        While Ir <> Selection.Rows.Count
                
            If ActiveSheet.Cells(startrow + Ir, startcol + Ic).Formula <> "" Then
                ActiveSheet.Cells(activerow, startcol).Formula = ActiveSheet.Cells(startrow + Ir, startcol + Ic).Formula
                activerow = activerow + 1
            End If
            
            Ir = Ir + 1
        Wend
        Ir = 0
        Ic = Ic + 1
    Wend
    
    ActiveSheet.Range(Cells(startrow, startcol + 1), Cells(endrow, endcol)).Value = ""


End Sub

Sub MultiRowToSingleRow()

    startrow = Selection.Row
    startcol = Selection.Column
    activecol = startcol
    endrow = Selection.Row + Selection.Rows.Count - 1
    endcol = Selection.Column + Selection.Columns.Count - 1
    
    Ic = 0
    Ir = 0
    
    While Ir <> Selection.Rows.Count
        While Ic <> Selection.Columns.Count
                
            If ActiveSheet.Cells(startrow + Ir, startcol + Ic).Formula <> "" Then
                ActiveSheet.Cells(startrow, activecol).Formula = ActiveSheet.Cells(startrow + Ir, startcol + Ic).Formula
                activecol = activecol + 1
            End If
            
            Ic = Ic + 1
        Wend
        Ic = 0
        Ir = Ir + 1
    Wend
    
    ActiveSheet.Range(Cells(startrow + 1, startcol), Cells(endrow, endcol)).Value = ""


End Sub



Sub HideUnusedRowsAndColumn()

Range(Columns(ActiveSheet.UsedRange.Columns.Count + 2), Columns(16384)).EntireColumn.Hidden = True
Rows(ActiveSheet.UsedRange.Rows.Count + 2 & ":1048576").EntireRow.Hidden = True
    
End Sub

Sub ShowUnusedRowsAndColumn()

Range(Columns(ActiveSheet.UsedRange.Columns.Count + 2), Columns(16384)).EntireColumn.Hidden = False
Rows(ActiveSheet.UsedRange.Rows.Count + 1 & ":1048576").EntireRow.Hidden = False
    
End Sub

Sub toggleColsAndRows()
    If Range(Columns(ActiveSheet.UsedRange.Columns.Count + 2), Columns(16384)).EntireColumn.Hidden Then
        Call Blad_Functies.ShowUnusedRowsAndColumn
    Else
        Call Blad_Functies.HideUnusedRowsAndColumn
    End If
End Sub


Sub hideColsAndRowsSelected()


Range(Columns(9), Columns(256)).EntireColumn.Hidden = True
    
        Range(Columns(Selection.Columns.Count + 1), Columns(16384)).EntireColumn.Hidden = True
        Rows(Selection.Rows.Count + 1 & ":1048576").EntireRow.Hidden = True
    
End Sub

Sub RemoveBars()

    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayGridlines = False
End Sub

Sub AddChart(Optional control As IRibbonControl)
    On Error Resume Next
    With ActiveSheet.Shapes.AddChart2(227, xlLine).Chart
        .SetSourceData Source:=Range(ActiveSheet.QueryTables(1).ResultRange.Address)
        
        .Axes(xlCategory).CategoryType = xlCategoryScale
        .Axes(xlCategory).ReversePlotOrder = True
    End With
End Sub

Sub AddChart2(Optional control As IRibbonControl)
    On Error Resume Next
    With ActiveSheet.Shapes.AddChart2(227, xlLine).Chart
        .SetSourceData Source:=Range(Selection.Address)
        
        .Axes(xlCategory).CategoryType = xlCategoryScale
        '.Axes(xlCategory).ReversePlotOrder = True
    End With
End Sub

Sub AddGaus(Optional control As IRibbonControl)
Call Grafiek_Invoegen.BuildGaus
End Sub


'String zoeken
Function ZoekenDeel(ZoekString, ZoekBereik As Range, Optional Helewaarde = False, Optional GeeftResterend = False)
Attribute ZoekenDeel.VB_Description = "Functie voor het zoeken naar een tekst in een range"
Attribute ZoekenDeel.VB_ProcData.VB_Invoke_Func = " \n25"
LookAtValue = xlPart
If Helewaarde Then
LookAtValue = xlWhole
End If
Set zoeker = ActiveSheet.Range(ZoekBereik.Address).Find(ZoekString, LookIn:=xlValues, LookAt:=LookAtValue)
If Not zoeker Is Nothing Then
If GeeftResterend Then
    ZoekenDeel = Replace(ActiveSheet.Range(zoeker.Address).Value, ZoekString, "")
Else
    ZoekenDeel = ActiveSheet.Range(zoeker.Address).Value
End If
Else
ZoekenDeel = "N/B"
End If
End Function

Function VanPuntNaarComma(BewerkString)
VanPuntNaarComma = Replace(BewerkString, ".", ",")
End Function

Sub toggleUnusedColRow(control As IRibbonControl)
    Call Blad_Functies.toggleColsAndRows
End Sub



Sub showShiftPersSelect(control As IRibbonControl)
getShiftPersonal.show

End Sub

Sub showCelOpmaak(control As IRibbonControl)
KeuzeCelOpmaak.show

End Sub

Sub showKKSInfo(control As IRibbonControl)
KKS_Info.show

End Sub

Sub showKKSSearch(control As IRibbonControl)
KKS_Zoeker.show

End Sub


Sub showDateTime(control As IRibbonControl)
DatePicker.show

End Sub

Sub showTabelInvoegen(control As IRibbonControl)
TabelInvoegen.show

End Sub


Sub showWeer(control As IRibbonControl)
Weer_Data.show

End Sub

Sub showTennet(control As IRibbonControl)
Tennet_Data.show
End Sub


Sub MergeExtra(control As IRibbonControl)
EndValue = ""
For Each cell In Selection
EndValue = EndValue & cell.Value & vbNewLine
Next
Application.DisplayAlerts = False
Selection.Merge
Application.DisplayAlerts = True
Selection.Value = EndValue
End Sub

Sub DeMergeExtra(control As IRibbonControl)

EndValueParts = Split(ActiveCell.Value, Chr(10))

Application.DisplayAlerts = False
Selection.UnMerge
Application.DisplayAlerts = True
Selection.Value = EndValueParts

End Sub

Sub SelectionToPDF(control As IRibbonControl)
With ActiveSheet.PageSetup
    .FitToPagesTall = False
    .FitToPagesWide = 1
    .Zoom = False
    .CenterHorizontally = True
    .RightFooter = "Created by JTools"
    With .LeftFooterPicture
        .Filename = "J:\Office\Grafisch\logo arn transparant.png"
        .Height = 60
    End With
    .LeftFooter = "&G"
    .CenterFooter = ActiveWorkbook.Name & " - " & Format(Now, "d-m-Y")
    
End With
TempPath = (Environ("Temp"))
Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:="" & TempPath & "\tmp_" & Format(Now(), "dmYhis") & ".pdf", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

Sub SendSelectionAsPDF(control As IRibbonControl)
With ActiveSheet.PageSetup
    .FitToPagesTall = False
    .FitToPagesWide = 1
    .Zoom = False
    .CenterHorizontally = True
    .RightFooter = "Created by JTools"
    With .LeftFooterPicture
        .Filename = "J:\Office\Grafisch\logo arn transparant.png"
        .Height = 60
    End With
    .LeftFooter = "&G"
    .CenterFooter = ActiveWorkbook.Name & " - " & Format(Now, "d-m-Y")
    
End With
TempPath = (Environ("Temp"))
TempFileName = TempPath & "\tmp_" & Format(Now(), "dmYhis") & ".pdf"
Selection.ExportAsFixedFormat Type:=xlTypePDF, Filename:=TempFileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

Dim objOutlook As Outlook.Application
Dim objMail As Outlook.MailItem
Set objOutlook = New Outlook.Application
Set objMail = objOutlook.CreateItem(olMailItem)

With objMail
    .Attachments.Add (TempFileName)
    .Display
    


End With


End Sub

Sub SendSelectionAsBody(control As IRibbonControl)
 Dim FSO As Object
 Dim ts As Object
Dim objOutlook As Outlook.Application
Dim objMail As Outlook.MailItem

If Selection Is Nothing Then
Exit Sub
End If

TempPath = "U:\JTOOLS" '(Environ("Temp"))

With ActiveWorkbook.PublishObjects.Add(xlSourceRange, TempPath & "tmp_selection.htm", ActiveSheet.Name, Selection.Address, xlHtmlStatic, ActiveWorkbook.Name, ActiveSheet.Name)
        .Publish (True)
        .AutoRepublish = False
    End With


 Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ts = FSO.GetFile(TempPath & "tmp_selection.htm").OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", "align=left x:publishsource=")
    RangetoHTML = Replace(RangetoHTML, "src=""tmp_selection_bestand\", "src=""" & TempPath & "tmp_selection_bestand\")
                                                    
 
Set objOutlook = New Outlook.Application
Set objMail = objOutlook.CreateItem(olMailItem)

With objMail
    .HTMLBody = RangetoHTML
    .Display
   
End With
On Error Resume Next
Kill TempPath & "tmp_selection.htm"
Kill TempPath & "tmp_selection_bestanden\*"
RmDir TempPath & "tmp_selection_bestanden\"
End Sub


