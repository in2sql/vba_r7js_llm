**Description / Описание:**
English: This code sets a sample text in cell B1, changes the font of characters 9-12 to "Font 1", and displays the font name in cell B3.
Russian: Этот код устанавливает пример текста в ячейку B1, изменяет шрифт символов с 9 по 12 на "Font 1" и отображает имя шрифта в ячейке B3.

```vba
' VBA Code to set text, modify font, and display font name

Sub ModifyFontExample()
    Dim ws As Worksheet
    Dim rng As Range
    Dim txt As String
    Dim fontName As String
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value in cell B1
    Set rng = ws.Range("B1")
    rng.Value = "This is just a sample text."
    
    ' Modify font of characters 9 to 12
    With rng.Characters(Start:=9, Length:=4).Font
        .Name = "Font 1"
        fontName = .Name
    End With
    
    ' Set the font name in cell B3
    ws.Range("B3").Value = "Font name: " & fontName
End Sub
```

```javascript
// JavaScript Code to set text, modify font, and display font name

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get range B1 and set its value
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get and set the font name to "Font 1"
var oFont = oCharacters.GetFont();
oFont.SetName("Font 1");

// Retrieve the font name
var sFontName = oFont.GetName();

// Set the font name in cell B3
oWorksheet.GetRange("B3").SetValue("Font name: " + sFontName);
```