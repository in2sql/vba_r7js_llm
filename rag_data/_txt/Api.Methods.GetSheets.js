# Description / Описание

This code adds a new sheet to the workbook, retrieves the collection of all sheets, obtains the names of the first two sheets, and sets these names into cells A1 and A2 of the second sheet.

Этот код добавляет новый лист в книгу, извлекает коллекцию всех листов, получает имена первых двух листов и записывает эти имена в ячейки A1 и A2 второго листа.

```vba
' Adds a new sheet, retrieves sheet names, and sets them in specific cells
Sub AddSheetAndSetValues()
    ' Add a new sheet named "new_sheet_name"
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "new_sheet_name"
    
    ' Get the collection of sheets
    Dim sheets As Sheets
    Set sheets = ThisWorkbook.Sheets
    
    ' Get names of the first two sheets
    Dim sheet_name1 As String
    Dim sheet_name2 As String
    sheet_name1 = sheets(1).Name
    sheet_name2 = sheets(2).Name
    
    ' Set the sheet names into cells A1 and A2 of the second sheet
    sheets(2).Range("A1").Value = sheet_name1
    sheets(2).Range("A2").Value = sheet_name2
End Sub
```

```javascript
// Adds a new sheet, retrieves sheet names, and sets them in specific cells
Api.AddSheet("new_sheet_name"); // Adds a new sheet named "new_sheet_name"

// Get the collection of all sheets
var sheets = Api.GetSheets();

// Get the names of the first two sheets
var sheet_name1 = sheets[0].GetName();
var sheet_name2 = sheets[1].GetName();

// Set the sheet names into cells A1 and A2 of the second sheet
sheets[1].GetRange("A1").SetValue(sheet_name1);
sheets[1].GetRange("A2").SetValue(sheet_name2);
```