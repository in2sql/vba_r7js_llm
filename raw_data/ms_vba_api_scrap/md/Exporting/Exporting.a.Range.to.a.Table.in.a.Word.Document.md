# Exporting a Range to a Table in a Word Document

## Business Description
This example takes the range A1:A10 on Sheet 1 and exports it to the first table in an existing Word document named 

## Behavior
This example takes the range A1:A10 on Sheet 1 and exports it to the first table in an existing Word document named "Table Report".

## Example Usage
```vba
Sub Export_Table_Data_Word()

    'Name of the existing Word document
    Const stWordDocument As String = "Table Report.docx"
    
    'Word objects.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim wdCell As Word.Cell
    
    'Excel objects
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    
    'Count used in a FOR loop to fill the Word table.
    Dim lnCountItems As Long
    
    'Variant to hold the data to be exported.
    Dim vaData As Variant
    
    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")
    vaData = wsSheet.Range("A1:A10").Value
    
    'Instantiate Word and open the "Table Reports" document.
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Open(wbBook.Path & "\" & stWordDocument)
    
    lnCountItems = 1
    
    'Place the data from the variant into the table in the Word doc.
    For Each wdCell In wdDoc.Tables(1).Columns(1).Cells
        wdCell.Range.Text = vaData(lnCountItems, 1)
        lnCountItems = lnCountItems + 1
    Next wdCell
    
    'Save and close the Word doc.
    With wdDoc
        .Save
        .Close
    End With
    
    wdApp.Quit
    
    'Null out the variables.
    Set wdCell = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    MsgBox "The " & stWordDocument & "'s table has succcessfully " & vbNewLine & _
           "been updated!", vbInformation

End Sub
```