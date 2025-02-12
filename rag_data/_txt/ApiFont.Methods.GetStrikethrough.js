**Description / Описание:**

*English:* This code sets a sample text in cell B1, applies strikethrough formatting to a specific part of the text, retrieves the strikethrough property, and displays its value in cell B3.

*Русский:* Этот код устанавливает пример текста в ячейку B1, применяет форматирование зачеркивания к определенной части текста, получает свойство зачеркивания и отображает его значение в ячейке B3.

```vba
' VBA Code

Sub SetStrikethrough()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Get the range B1 and set its value
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    oRange.Value = "This is just a sample text."

    ' Apply strikethrough to characters 9 to 12
    With oRange.Characters(Start:=9, Length:=4).Font
        .Strikethrough = True
    End With

    ' Get the strikethrough property
    Dim bStrikethrough As Boolean
    bStrikethrough = oRange.Characters(Start:=9, Length:=4).Font.Strikethrough

    ' Set the value in B3
    oWorksheet.Range("B3").Value = "Strikethrough property: " & bStrikethrough
End Sub
```

```javascript
// JavaScript Code

// This example shows how to get the strikethrough property of the specified font.
// Этот пример показывает, как получить свойство зачеркивания указанного шрифта.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set the value
var oCharacters = oRange.GetCharacters(9, 4); // Get characters from position 9, length 4
var oFont = oCharacters.GetFont(); // Get the font of the characters
oFont.SetStrikethrough(true); // Apply strikethrough
var bStrikethrough = oFont.GetStrikethrough(); // Get the strikethrough property
oWorksheet.GetRange("B3").SetValue("Strikethrough property: " + bStrikethrough); // Set value in B3
```