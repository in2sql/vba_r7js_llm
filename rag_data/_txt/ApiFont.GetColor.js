**Description:**
This code sets the text in cell B1 to "This is just a sample text." It then changes the font color of the word "just" to a specific RGB color and applies the same color to the word "sample".

Этот код устанавливает текст в ячейке B1 как "This is just a sample text.". Затем он изменяет цвет шрифта слова "just" на определенный RGB цвет и применяет тот же цвет к слову "sample".

```vba
' VBA Code
Sub SetFontColors()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters
    Dim oFont As Font
    Dim oColor As Long
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get range B1
    Set oRange = oWorksheet.Range("B1")
    
    ' Set value in B1
    oRange.Value = "This is just a sample text."
    
    ' Get characters from position 9, length 4 ("just")
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    ' Get the font of these characters
    Set oFont = oCharacters.Font
    
    ' Create color from RGB(255, 111, 61)
    oColor = RGB(255, 111, 61)
    
    ' Set font color
    oFont.Color = oColor
    
    ' Get the font color
    oColor = oFont.Color
    
    ' Get characters from position 16, length 6 ("sample")
    Set oCharacters = oRange.Characters(Start:=16, Length:=6)
    
    ' Get the font of these characters
    Set oFont = oCharacters.Font
    
    ' Set font color
    oFont.Color = oColor
End Sub
```

```javascript
// JavaScript Code
// This script sets the text in cell B1 and changes the font color of specific words

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get range B1
var oRange = oWorksheet.GetRange("B1");

// Set value in B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9, length 4 ("just")
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of these characters
var oFont = oCharacters.GetFont();

// Create color from RGB(255, 111, 61)
var oColor = Api.CreateColorFromRGB(255, 111, 61);

// Set font color
oFont.SetColor(oColor);

// Get the font color
oColor = oFont.GetColor();

// Get characters from position 16, length 6 ("sample")
oCharacters = oRange.GetCharacters(16, 6);

// Get the font of these characters
oFont = oCharacters.GetFont();

// Set font color
oFont.SetColor(oColor);
```