Sub setStartPage(page As Integer)
'
' Macro1 Macro
'
'
    If ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        ActiveWindow.Panes(2).Close
    End If
    If ActiveWindow.ActivePane.View.Type = wdNormalView Or ActiveWindow. _
        ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    With Selection.HeaderFooter.PageNumbers
        .NumberStyle = wdPageNumberStyleArabic
        .HeadingLevelForChapter = 0
        .IncludeChapterNumber = False
        .ChapterPageSeparator = wdSeparatorHyphen
        .RestartNumberingAtSection = True
        .StartingNumber = page
    End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

Sub Paste()

  Dim objExcel As Object
  Dim exWb As Object

  Set objExcel = CreateObject("Excel.Application")

        Set exWb = objExcel.Workbooks.Open("Excel file path")

  Dim i As Integer
  Dim j As Integer
  j = 1

  For i = 1 To 347
              
    If i Mod 21 = 0 Then
        setStartPage (i)
        j = 1
    End If

      Selection.GoTo wdGoToPage, wdGoToAbsolute, j
      Selection.Paragraphs(1).Range.Delete
    
      Selection.Font.Bold = True
      Selection.Font.Underline = True
      Selection.ParagraphFormat.SpaceAfter = 0
      Selection.ParagraphFormat.SpaceBefore = 12
      Selection.TypeText Text:=exWb.Sheets("Sheet1").Cells(i, 2) & " " & exWb.Sheets("Sheet1").Cells(i, 3)
      Selection.TypeText Text:=MyText
      Selection.InsertParagraph
      
      
      
      ActiveDocument.ExportAsFixedFormat OutputFileName:= _
            ("path to save files\" & exWb.Sheets("Sheet1").Cells(i, 2) & ".pdf"), _
            ExportFormat:=wdExportFormatPDF, _
            Range:=wdExportFromTo, From:=j, To:=j

    j = j + 1
  Next i


  exWb.Close

  Set exWb = Nothing

End Sub
