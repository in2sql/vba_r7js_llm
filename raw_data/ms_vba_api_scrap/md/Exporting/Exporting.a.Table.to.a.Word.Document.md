# Exporting a Table to a Word Document

## Business Description
This example takes the table named 

## Behavior
This example takes the table named "Table1" on Sheet 1 and copies it into an existing Word document named "Quarter Report" at the bookmarked location named "Report".

## Example Usage
```vba
Sub Export_Table_Word()

    'Name of the existing Word doc.
    Const stWordReport As String = "Quarter Report.docx"
    
    'Word objects.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim wdbmRange As Word.Range
    
    'Excel objects.
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim rnReport As Range
    
    'Initialize the Excel objects.
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")
    Set rnReport = wsSheet.Range("Table1")
    
    'Initialize the Word objets.
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Open(wbBook.Path & "\" & stWordReport)
    Set wdbmRange = wdDoc.Bookmarks("Report").Range
    
    'If the macro has been run before, clean up any artifacts before trying to paste the table in again.
    On Error Resume Next
    With wdDoc.InlineShapes(1)
        .Select
        .Delete
    End With
    On Error GoTo 0
    
    'Turn off screen updating.
    Application.ScreenUpdating = False
    
    'Copy the report to the clipboard.
    rnReport.Copy
    
    'Select the range defined by the "Report" bookmark and paste in the report from clipboard.
    With wdbmRange
        .Select
        .PasteSpecial Link:=False, _
                      DataType:=wdPasteMetafilePicture, _
                      Placement:=wdInLine, _
                      DisplayAsIcon:=False
    End With
    
    'Save and close the Word doc.
    With wdDoc
        .Save
        .Close
    End With
    
    'Quit Word.
    wdApp.Quit
    
    'Null out your variables.
    Set wdbmRange = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    'Clear out the clipboard, and turn screen updating back on.
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
    End With
    
    MsgBox "The report has successfully been " & vbNewLine & _
           "transferred to " & stWordReport, vbInformation

End Sub
```