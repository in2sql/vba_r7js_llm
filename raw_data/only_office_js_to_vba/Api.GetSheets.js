**English:** This code demonstrates how to add a new sheet, access all sheets, retrieve sheet names, and set values in specific cells of a sheet.

**Russian:** Этот код демонстрирует, как добавить новый лист, получить доступ ко всем листам, получить имена листов и установить значения в определенные ячейки листа.

```vba
' This subroutine adds a new sheet, retrieves the sheets collection,
' gets the names of the first two sheets, and sets values in cells A1 and A2 of the second sheet.
Sub Example()
    ' Add a new sheet named "new_sheet_name"
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "new_sheet_name"
    
    ' Get the collection of all sheets
    Dim sheetsCollection As Sheets
    Set sheetsCollection = ThisWorkbook.Sheets
    
    ' Get the names of the first and second sheets
    Dim sheet_name1 As String
    Dim sheet_name2 As String
    sheet_name1 = sheetsCollection(1).Name
    sheet_name2 = sheetsCollection(2).Name
    
    ' Set the value of cell A1 in the second sheet to sheet_name1
    sheetsCollection(2).Range("A1").Value = sheet_name1
    
    ' Set the value of cell A2 in the second sheet to sheet_name2
    sheetsCollection(2).Range("A2").Value = sheet_name2
End Sub
```

```javascript
// This script adds a new sheet, retrieves the sheets collection,
// gets the names of the first two sheets, and sets values in cells A1 and A2 of the second sheet.
Api.AddSheet("new_sheet_name"); // Add a new sheet named "new_sheet_name"
var sheets = Api.GetSheets(); // Get the collection of all sheets
var sheet_name1 = sheets[0].GetName(); // Get the name of the first sheet
var sheet_name2 = sheets[1].GetName(); // Get the name of the second sheet
sheets[1].GetRange("A1").SetValue(sheet_name1); // Set cell A1 in the second sheet to sheet_name1
sheets[1].GetRange("A2").SetValue(sheet_name2); // Set cell A2 in the second sheet to sheet_name2
```