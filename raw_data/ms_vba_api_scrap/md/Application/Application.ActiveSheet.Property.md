# Application.ActiveSheet property (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Syntax
```vba
MsgBox "The name of the active sheet is " & ActiveSheet.Name
```

## Remarks
If you don't specify an object qualifier, this property returns the active sheet in the active workbook.

## Example
```vba
Sub PrintSheets()

   'Set up your variables.
   Dim iRow As Integer, iRowL As Integer, iPage As Integer
   'Find the last row that contains data.
   iRowL = Cells(Rows.Count, 1).End(xlUp).Row
   
   'Define the print area as the range containing all the data in the first two columns of the current worksheet.
   ActiveSheet.PageSetup.PrintArea = Range("A1:B" & iRowL).Address
   
   'Select all the rows containing data.
   Rows(iRowL).Select
   
   'display the automatic page breaks
   ActiveSheet.DisplayAutomaticPageBreaks = True
   Range("B1").Value = "Page 1"
   
   'After each page break, go to the next cell in column B and write out the page number.
   For iPage = 1 To ActiveSheet.HPageBreaks.Count
      ActiveSheet.HPageBreaks(iPage) _
         .Location.Offset(0, 1).Value = "Page " & iPage + 1
   Next iPage
   
   'Show the print preview, and afterwards remove the page numbers from column B.
   ActiveSheet.PrintPreview
   Columns("B").ClearContents
   Range("A1").Select
End Sub
```

