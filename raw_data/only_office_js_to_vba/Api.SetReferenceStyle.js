### Description / Описание

This example sets the reference style to R1C1 and writes the current reference style into cell A1.

Этот пример устанавливает стиль ссылок на R1C1 и записывает текущий стиль ссылок в ячейку A1.

```vba
' This example sets the reference style to R1C1 and writes the current reference style into cell A1
Sub SetReferenceStyle()
    ' Set reference style to R1C1
    Application.ReferenceStyle = xlR1C1
    ' Write the current reference style to cell A1
    Range("A1").Value = Application.ReferenceStyle
End Sub
```

```javascript
// This example sets the reference style to R1C1 and writes the current reference style into cell A1
function setReferenceStyle() {
    // Get the active sheet
    var oWorksheet = Api.GetActiveSheet();
    // Set reference style to R1C1
    Api.SetReferenceStyle("xlR1C1");
    // Set value of cell A1 to the current reference style
    oWorksheet.GetRange("A1").SetValue(Api.GetReferenceStyle());
}
```