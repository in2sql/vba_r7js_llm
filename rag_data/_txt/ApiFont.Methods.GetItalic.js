### Description

This code sets a sample text in cell B1, applies italic formatting to a specific part of the text, and displays the italic property status in cell B3.

Этот код устанавливает пример текста в ячейку B1, применяет курсивное форматирование к определенной части текста и отображает статус свойства курсивного шрифта в ячейке B3.

```javascript
// This example shows how to get the italic property of the specified font.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
// Set the value of cell B1
oRange.SetValue("This is just a sample text.");
// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);
var oFont = oCharacters.GetFont();
// Set italic to true
oFont.SetItalic(true);
// Get the italic property
var bItalic = oFont.GetItalic();
// Display the italic property in cell B3
oWorksheet.GetRange("B3").SetValue("Italic property: " + bItalic);
```

```vba
' This VBA code sets a sample text in cell B1, applies italic formatting to specific characters, and displays the italic property status in cell B3.

Sub SetItalicProperty()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Set value in cell B1
    ws.Range("B1").Value = "This is just a sample text."
    
    ' Apply italic to characters 9 to 12 ("just")
    With ws.Range("B1").Characters(Start:=9, Length:=4).Font
        .Italic = True
    End With
    
    ' Get the italic property
    Dim isItalic As Boolean
    isItalic = ws.Range("B1").Characters(Start:=9, Length:=4).Font.Italic
    
    ' Set value in cell B3
    ws.Range("B3").Value = "Italic property: " & isItalic
End Sub
```