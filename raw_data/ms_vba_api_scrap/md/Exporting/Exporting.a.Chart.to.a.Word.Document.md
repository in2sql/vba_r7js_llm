# Exporting a Chart to a Word Document

## Business Description
This example takes a chart named 

## Behavior
This example takes a chart named "Chart 1" from Sheet 1 and exports it as a .gif file. Then it inserts the .gif file into an existing Word document named "Chart Report" at the bookmarked location named "ChartReport".

## Example Usage
```vba
Sub Export_Chart_Word()

    'Name of an existing Word document, and the name the chart will have when exported.
    Const stWordDocument As String = "Chart Report.docx"
    Const stChartName As String = "ChartReport.gif"
    
    'Word objects.
    Dim wdApp As Word.Application
    Dim wdDoc As Word.Document
    Dim wdbmRange As Word.Range
    
    'Excel objects.
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim ChartObj As ChartObject
    
    'Initialize the Excel objets.
    Set wbBook = ThisWorkbook
    Set wsSheet = wbBook.Worksheets("Sheet1")
    Set ChartObj = wsSheet.ChartObjects("Chart 1")
    
    'Turn off screen updating.
    Application.ScreenUpdating = False
    
    'Export the chart to the current directory, using the specified name, and save the chart as a .gif
    ChartObj.Chart.Export _
                   Filename:=wbBook.Path & "\" & stChartName, _
                   FilterName:="GIF"
    
    'Initialize the Word objets to the existing Word document and bookmark.
    Set wdApp = New Word.Application
    Set wdDoc = wdApp.Documents.Open(wbBook.Path & "\" & stWordDocument)
    Set wdbmRange = wdDoc.Bookmarks("ChartReport").Range
    
    'If there is already an inline shape, that means the macro has been run before - clean up any artifacts.
    On Error Resume Next
    With wdDoc.InlineShapes(1)
        .Select
        .Delete
    End With
    On Error GoTo 0
    
    'Add the .gif file to the document at the bookmarked location,
    'and ensure that it is saved inside the Word doc.
    With wdbmRange
        .Select
        .InlineShapes.AddPicture _
        Filename:=wbBook.Path & "\" & stChartName, _
        LinkToFile:=False, _
        savewithdocument:=True
    End With
    
    'Save and close the Word document.
    With wdDoc
        .Save
        .Close
    End With
    
    'Quit Word.
    wdApp.Quit
    
    'Clear the variables.
    Set wdbmRange = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    'Delete the temporary .gif file.
    On Error Resume Next
    Kill wbBook.Path & "\" & stChartName
    On Error GoTo 0
    
    MsgBox "Chart exported successfully to " & stWordDocument

End Sub
```