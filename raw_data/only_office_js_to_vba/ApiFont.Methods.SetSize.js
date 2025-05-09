**Description:**
Sets the font size of specific characters in cell B1 to 18.
Устанавливает размер шрифта для определенных символов в ячейке B1 на 18.

```vba
' This example sets the font size property to the specified font.
Sub SetFontSize()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    oRange.Value = "This is just a sample text."
    
    Dim oCharacters As Characters
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    Dim oFont As Font
    Set oFont = oCharacters.Font
    oFont.Size = 18
End Sub
```

```javascript
// This example sets the font size property to the specified font.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(9, 4);
var oFont = oCharacters.GetFont();
oFont.SetSize(18);
```