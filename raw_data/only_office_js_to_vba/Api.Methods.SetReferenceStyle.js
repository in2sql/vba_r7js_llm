**English:** This example sets the reference style to R1C1 and writes the current reference style to cell A1.  
**Russian:** Этот пример устанавливает стиль ссылок R1C1 и записывает текущий стиль ссылок в ячейку A1.

```vba
' This example sets the reference style to R1C1 and writes the current reference style to cell A1
' Этот пример устанавливает стиль ссылок R1C1 и записывает текущий стиль ссылок в ячейку A1

Sub SetReferenceStyle()
    ' Get the active sheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set the reference style to R1C1
    Application.ReferenceStyle = xlR1C1
    
    ' Write the current reference style to cell A1
    oWorksheet.Range("A1").Value = Application.ReferenceStyle
End Sub
```

```javascript
// This example sets the reference style to R1C1 and writes the current reference style to cell A1
// Этот пример устанавливает стиль ссылок R1C1 и записывает текущий стиль ссылок в ячейку A1

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set the reference style to R1C1
Api.SetReferenceStyle("xlR1C1");

// Write the current reference style to cell A1
oWorksheet.GetRange("A1").SetValue(Api.GetReferenceStyle());
```