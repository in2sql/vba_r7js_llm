**Description / Описание**

This code sets the bold property to the specified font.
Этот код устанавливает свойство жирного начертания для указанного шрифта.

```vba
' This example sets part of the text in cell B1 to bold.
' Этот пример устанавливает часть текста в ячейке B1 жирным шрифтом.

Sub SetBoldText()
    Dim oRange As Range
    ' Set the range to cell B1
    Set oRange = ActiveSheet.Range("B1")
    ' Set the text value
    oRange.Value = "This is just a sample text."
    ' Set characters 9 to 12 to bold
    oRange.Characters(Start:=9, Length:=4).Font.Bold = True
End Sub
```

```javascript
// This example sets the bold property to the specified font.
// Этот код устанавливает свойство жирного начертания для указанного шрифта.

var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set cell value
var oCharacters = oRange.GetCharacters(9, 4); // Get characters 9 to 12
var oFont = oCharacters.GetFont(); // Get the font of specified characters
oFont.SetBold(true); // Set bold to true
```