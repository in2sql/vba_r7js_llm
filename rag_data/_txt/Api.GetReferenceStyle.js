## Description / Описание

**English:** This example gets the reference style and sets it to cell A1.

**Russian:** Этот пример получает стиль ссылок и устанавливает его в ячейку A1.

```vba
' This example gets the reference style
Sub SetReferenceStyle()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    ' Set the value of cell A1 to the current reference style
    oWorksheet.Range("A1").Value = Application.ReferenceStyle
End Sub
```

```javascript
// This example gets the reference style
var oWorksheet = Api.GetActiveSheet();
// Set the value of cell A1 to the current reference style
oWorksheet.GetRange("A1").SetValue(Api.GetReferenceStyle());
```