```
// This example saves changes to the specified document.
// Этот пример сохраняет изменения в указанном документе.
```

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the value of cell A1
oWorksheet.GetRange("A1").SetValue("This sample text is saved to the worksheet.");

// Save the document
Api.Save();
```

```vba
' Получение активного рабочего листа
Dim oWorksheet As Worksheet
Set oWorksheet = Api.GetActiveSheet()

' Установка значения ячейки A1
oWorksheet.Range("A1").Value = "This sample text is saved to the worksheet."

' Сохранение документа
Api.Save
```