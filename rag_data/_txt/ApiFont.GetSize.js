# Example: Get and Set Font Size Properties / Пример: Получение и установка свойств размера шрифта

This example demonstrates how to get the font size property of a specified font, set a new font size, and display the updated size.

Этот пример демонстрирует, как получить свойство размера шрифта указанного шрифта, установить новый размер шрифта и отобразить обновленный размер.

```javascript
// This example shows how to get the font size property of the specified font.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get range B1
var oRange = oWorksheet.GetRange("B1");

// Set a value in B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the specified characters
var oFont = oCharacters.GetFont();

// Set the font size to 18
oFont.SetSize(18);

// Get the font size
var nSize = oFont.GetSize();

// Set the font size value in cell B3
oWorksheet.GetRange("B3").SetValue("Size property: " + nSize);
```

```vba
' This example demonstrates how to get the font size property of a specified font,
' set a new font size, and display the updated size.

Sub GetSetFontSize()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Get range B1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    
    ' Set a value in B1
    oRange.Value = "This is just a sample text."
    
    ' Get characters from position 9 with length 4
    Dim oCharacters As Characters
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    ' Set the font size to 18
    oCharacters.Font.Size = 18
    
    ' Get the font size
    Dim nSize As Double
    nSize = oCharacters.Font.Size
    
    ' Set the font size value in cell B3
    oWorksheet.Range("B3").Value = "Size property: " & nSize
End Sub
```