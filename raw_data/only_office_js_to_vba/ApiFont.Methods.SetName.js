**Description / Описание**
- **English:** This code sets the font name of specific characters in cell B1 and displays the font name in cell B3.
- **Russian:** Этот код устанавливает имя шрифта для определённых символов в ячейке B1 и отображает имя шрифта в ячейке B3.

```javascript
// This example sets the font name property to the specified font.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(9, 4);
var oFont = oCharacters.GetFont();
oFont.SetName("Font 1");
var sFontName = oFont.GetName();
oWorksheet.GetRange("B3").SetValue("Font name: " + sFontName); 
```

```vba
' This example sets the font name property to the specified font.
Sub SetFontName()
    Dim oSheet As Worksheet
    Set oSheet = ThisWorkbook.ActiveSheet
    
    Dim oRange As Range
    Set oRange = oSheet.Range("B1")
    
    oRange.Value = "This is just a sample text."
    
    Dim oCharacters As Characters
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    With oCharacters.Font
        .Name = "Font 1"
    End With
    
    Dim sFontName As String
    sFontName = oCharacters.Font.Name
    
    oSheet.Range("B3").Value = "Font name: " & sFontName
End Sub
```