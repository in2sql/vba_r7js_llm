# Description / Описание

**English:** This code gets the active worksheet and sets the value of cell A1 to the current reference style.

**Russian:** Этот код получает активный лист и устанавливает значение ячейки A1 в текущий стиль ссылок.

```vba
' This VBA code gets the active worksheet and sets the value of cell A1 to the current reference style.
Sub SetReferenceStyle()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    oWorksheet.Range("A1").Value = Application.ReferenceStyle
End Sub
```

```js
// This example gets the active sheet and sets the value of cell A1 to the current reference style.
var oWorksheet = Api.GetActiveSheet();
oWorksheet.GetRange("A1").SetValue(Api.GetReferenceStyle());
```