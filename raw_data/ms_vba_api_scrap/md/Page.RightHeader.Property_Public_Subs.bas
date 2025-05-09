Attribute VB_Name = "Public_Subs"
Option Explicit


Sub Optimize()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

End Sub


Sub De_Optimize()

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub Optimize_with_Calc()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual

End Sub

Sub De_Optimize_with_Calc()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationAutomatic

End Sub


Sub Print_PDF()

Dim StartRng As Range
Dim LastRow As Long
Dim PDFArea As Range
Dim RAWRetn As VbMsgBoxResult

'Unhiding Sheet
ShPDF.Visible = xlSheetVisible

Set StartRng = ShPDF.Range("B2:M2")
LastRow = ShPDF.Range("B" & Rows.Count).End(xlUp).Row
Set PDFArea = ShPDF.Range(StartRng, ShPDF.Range("M" & LastRow))

'Debug.Print PDFArea.Address
    
    'Printing to PDF
    Application.PrintCommunication = True
    Application.PrintCommunication = False
        
        With ShPDF.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = xlPortrait
            .Draft = False
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = xlPrintErrorsBlank
            .OddAndEvenPagesHeaderFooter = False
            .DifferentFirstPageHeaderFooter = False
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
            .EvenPage.LeftHeader.Text = ""
            .EvenPage.CenterHeader.Text = ""
            .EvenPage.RightHeader.Text = ""
            .EvenPage.LeftFooter.Text = ""
            .EvenPage.CenterFooter.Text = ""
            .EvenPage.RightFooter.Text = ""
            .FirstPage.LeftHeader.Text = ""
            .FirstPage.CenterHeader.Text = ""
            .FirstPage.RightHeader.Text = ""
            .FirstPage.LeftFooter.Text = ""
            .FirstPage.CenterFooter.Text = ""
            .FirstPage.RightFooter.Text = ""
        End With
        
    Application.PrintCommunication = True
    PDFArea.PrintOut Copies:=1, Collate:=True


'Hiding Sheet, Toggling back Dropdown trigger & Unprotecting Sheet - ShPDF
With ShPDF
.Visible = xlSheetHidden
.Unprotect
.Range("XFD1") = 1
End With

'Disabling Preview button and Clearinig 'Read-for-download' Message text
With ShImport
.Range("L20").ClearContents
.BtPreview.Enabled = False
End With

'Success Message & Raw Data Retention offer
RAWRetn = MsgBox("Your PDF has been saved succesfully." & vbNewLine & "Retain RAW DATA to Download Next Marksheet?", _
vbYesNo, "Success!")

If RAWRetn = vbYes Then

    Exit Sub
    
ElseIf RAWRetn = vbNo Then

    Application.DisplayAlerts = False
    
    ThisWorkbook.Sheets("RAW DATA_Imported").Delete
    
    Application.DisplayAlerts = True
    
    ShPDF.Range("XEU1").Value = 0 'Toggling Trigger for Download Button (BtDwnld)
    ShImport.CbID.Value = "< Select Student ID >"

End If


End Sub

Sub Delete_Previous_RAWdata()

Dim PrevSheet As Worksheet

On Error Resume Next

Call Optimize

Set PrevSheet = ThisWorkbook.Sheets("RAW DATA_Imported")

'Check for Previous RAW DATA sheet and Delete it
If Not PrevSheet Is Nothing Then

PrevSheet.Delete

End If

Call De_Optimize

End Sub
