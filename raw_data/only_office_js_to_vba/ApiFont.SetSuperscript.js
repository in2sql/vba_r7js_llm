```plaintext
// Description in English
This example sets the superscript property to the specified font.

// Описание на русском
Этот пример устанавливает свойство надстрочного знака для указанного шрифта.
```

```vba
' VBA Code
Sub SetSuperscript()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Get the range B1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")

    ' Set the value of the range
    oRange.Value = "This is just a sample text."

    ' Get the characters from position 9 with length 4
    Dim oCharacters As Characters
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)

    ' Set superscript to True
    oCharacters.Font.Superscript = True
End Sub
```

```javascript
// OnlyOffice JS Code
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of the range
oRange.SetValue("This is just a sample text.");

// Get the characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the characters
var oFont = oCharacters.GetFont();

// Set superscript to true
oFont.SetSuperscript(true);
```