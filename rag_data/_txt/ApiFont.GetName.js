#### This example shows how to get the font name property of the specified font.
#### Этот пример демонстрирует, как получить свойство имени шрифта указанного шрифта.

```vba
' VBA code to get and set font name in Excel

Sub GetAndSetFontName()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters
    Dim oFont As Font
    Dim sFontName As String
    
    Set oWorksheet = ThisWorkbook.ActiveSheet ' Get active sheet
    Set oRange = oWorksheet.Range("B1") ' Get range B1
    oRange.Value = "This is just a sample text." ' Set value in B1
    Set oCharacters = oRange.Characters(Start:=9, Length:=4) ' Get characters from position 9 to 12
    Set oFont = oCharacters.Font ' Get font of the characters
    oFont.Name = "Font 1" ' Set font name to "Font 1"
    sFontName = oFont.Name ' Get font name
    oWorksheet.Range("B3").Value = "Font name: " & sFontName ' Set value in B3
End Sub
```

```javascript
// OnlyOffice JS code to get and set font name

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value in B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9, length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the characters
var oFont = oCharacters.GetFont();

// Set the font name to "Font 1"
oFont.SetName("Font 1");

// Get the font name
var sFontName = oFont.GetName();

// Set the value in B3
oWorksheet.GetRange("B3").SetValue("Font name: " + sFontName);
```