# Setting Italic Property to a Specified Font  
# Установка свойства курсив к указанному шрифту

**VBA Code:**

```vba
' This macro sets the italic property to the specified font in cell B1
Sub SetItalicFont()
    Dim ws As Worksheet
    Dim rng As Range
    Dim chr As Characters
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Define the range B1
    Set rng = ws.Range("B1")
    
    ' Set the cell value
    rng.Value = "This is just a sample text."
    
    ' Get characters from position 9 with length 4
    Set chr = rng.Characters(Start:=9, Length:=4)
    
    ' Set the font to italic
    chr.Font.Italic = True
End Sub
```

**OnlyOffice JS Code:**

```javascript
// This example sets the italic property to the specified font.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(9, 4);
var oFont = oCharacters.GetFont();
oFont.SetItalic(true);
```