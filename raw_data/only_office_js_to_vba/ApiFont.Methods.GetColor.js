**Description / Описание**

This code sets specific portions of text in cell B1 to a defined font color.

Этот код устанавливает определенные части текста в ячейке B1 на заданный цвет шрифта.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 to 12
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Create a color with RGB values
var oColor = Api.CreateColorFromRGB(255, 111, 61);

// Set the font color of the selected characters
oFont.SetColor(oColor);

// Retrieve the font color
oColor = oFont.GetColor();

// Get characters from position 16 to 21
oCharacters = oRange.GetCharacters(16, 6);

// Get the font of the new selection
oFont = oCharacters.GetFont();

// Set the font color of the new selection to the previously retrieved color
oFont.SetColor(oColor);
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Get the range B1
Dim oRange As Range
Set oRange = oWorksheet.Range("B1")

' Set the value of B1
oRange.Value = "This is just a sample text."

' Define the start position and length for the first set of characters
Dim startPos1 As Integer
Dim length1 As Integer
startPos1 = 9
length1 = 4

' Get the first substring
Dim firstSubstring As String
firstSubstring = Mid(oRange.Value, startPos1, length1)

' Define the color using RGB
Dim oColor As Long
oColor = RGB(255, 111, 61)

' Apply the color to the first substring
With oRange.Characters(startPos1, length1).Font
    .Color = oColor
End With

' Define the start position and length for the second set of characters
Dim startPos2 As Integer
Dim length2 As Integer
startPos2 = 16
length2 = 6

' Apply the previously defined color to the second substring
With oRange.Characters(startPos2, length2).Font
    .Color = oColor
End With
```