```javascript
// Description:
// This code retrieves the "Sheet1" worksheet and sets the value of cell A1 to a sample text.
// Описание:
// Этот код получает лист "Sheet1" и устанавливает значение ячейки A1 на пример текста.

// JavaScript (OnlyOffice API) code
// Get the sheet named "Sheet1"
var oWorksheet = Api.GetSheet("Sheet1");
// Set the value of cell A1
oWorksheet.GetRange("A1").SetValue("This is a sample text on 'Sheet1'.");
```

```vba
' Description:
' This code retrieves the "Sheet1" worksheet and sets the value of cell A1 to a sample text.
' Описание:
' Этот код получает лист "Sheet1" и устанавливает значение ячейки A1 на пример текста.

' VBA code
' Get the sheet named "Sheet1"
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.Sheets("Sheet1")
' Set the value of cell A1
oWorksheet.Range("A1").Value = "This is a sample text on 'Sheet1'."
```