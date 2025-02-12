```plaintext
// Description in English: This script sets the strikethrough property for specific characters in cell B1.
// Описание на русском: Этот скрипт устанавливает свойство зачеркивания для определенных символов в ячейке B1.

````vba
' This VBA script sets the strikethrough property for specific characters in cell B1
Sub SetStrikethrough()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    With ws.Range("B1")
        .Value = "This is just a sample text."
        ' Applying strikethrough to characters 9 to 12
        .Characters(Start:=9, Length:=4).Font.Strikethrough = True
    End With
End Sub
````

```javascript
// This JavaScript code sets the strikethrough property for specific characters in cell B1
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(9, 4); // Characters 9 to 12
var oFont = oCharacters.GetFont();
oFont.SetStrikethrough(true);
```