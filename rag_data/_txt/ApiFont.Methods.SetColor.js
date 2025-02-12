### Description / Описание

**English:** This code sets the font color of a specific portion of text in cell B1 to a specified RGB color.

**Русский:** Этот код устанавливает цвет шрифта определенной части текста в ячейке B1 на заданный RGB цвет.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of cell B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with a length of 4 characters
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Create a color from RGB values
var oColor = Api.CreateColorFromRGB(255, 111, 61);

// Set the font color
oFont.SetColor(oColor);
```

```vba
' Get the active worksheet
Dim oWorksheet As Object
Set oWorksheet = Api.GetActiveSheet()

' Get the range B1
Dim oRange As Object
Set oRange = oWorksheet.GetRange("B1")

' Set the value of cell B1
oRange.SetValue "This is just a sample text."

' Get characters from position 9 with a length of 4 characters
Dim oCharacters As Object
Set oCharacters = oRange.GetCharacters(9, 4)

' Get the font of the selected characters
Dim oFont As Object
Set oFont = oCharacters.GetFont()

' Create a color from RGB values
Dim oColor As Object
Set oColor = Api.CreateColorFromRGB(255, 111, 61)

' Set the font color
oFont.SetColor oColor
```