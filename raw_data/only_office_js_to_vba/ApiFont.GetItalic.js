**English:**  
This code sets a value in cell B1, applies italic formatting to characters 9-12, and writes the italic property state in cell B3.

**Russian:**  
Этот код устанавливает значение в ячейке B1, применяет курсивное форматирование к символам с 9 по 12 и записывает состояние свойства курсива в ячейку B3.

```vba
' VBA Code to manipulate cell formatting and values

Sub SetItalicProperty()
    Dim ws As Worksheet
    Dim rng As Range
    Dim chars As Characters
    Dim isItalic As Boolean
    
    ' Set the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set value in cell B1
    Set rng = ws.Range("B1")
    rng.Value = "This is just a sample text."
    
    ' Get characters from position 9 with length 4
    Set chars = rng.Characters(Start:=9, Length:=4)
    
    ' Set italic to True
    chars.Font.Italic = True
    
    ' Get the italic property
    isItalic = chars.Font.Italic
    
    ' Write the italic property state to cell B3
    ws.Range("B3").Value = "Italic property: " & isItalic
End Sub
```

```javascript
// JavaScript Code to manipulate cell formatting and values using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value in cell B1
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Set italic to true
oFont.SetItalic(true);

// Get the italic property
var bItalic = oFont.GetItalic();

// Write the italic property state to cell B3
oWorksheet.GetRange("B3").SetValue("Italic property: " + bItalic);
```