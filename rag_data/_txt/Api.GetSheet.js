**Description:**
This code retrieves the "Sheet1" worksheet and sets the value of cell A1 to a sample text.

Этот код получает рабочий лист "Sheet1" и устанавливает значение ячейки A1 в пример текста.

```javascript
// This example shows how to get an object that represents a sheet.
var oWorksheet = Api.GetSheet("Sheet1");
// Set the value of cell A1 to a sample text on 'Sheet1'.
oWorksheet.GetRange("A1").SetValue("This is a sample text on 'Sheet1'.");
```

```vba
' This VBA code retrieves the "Sheet1" worksheet and sets the value of cell A1 to a sample text.
Sub SetSampleText()
    Dim oWorksheet As Worksheet
    ' Get the worksheet named "Sheet1"
    Set oWorksheet = ThisWorkbook.Sheets("Sheet1")
    ' Set the value of cell A1
    oWorksheet.Range("A1").Value = "This is a sample text on 'Sheet1'."
End Sub
```