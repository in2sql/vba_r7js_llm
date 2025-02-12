**English & Russian Description**

This example sets the font size property to the specified font in cell B1.
Этот пример устанавливает размер шрифта для заданного шрифта в ячейке B1.

```vba
' VBA Code to set font size in cell B1

Sub SetFontSize()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters
    Dim oFont As Font
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the value of B1
    oRange.Value = "This is just a sample text."
    
    ' Get characters from position 9, length 4
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    ' Get the font of the specified characters
    Set oFont = oCharacters.Font
    
    ' Set the font size to 18
    oFont.Size = 18
End Sub
```

```javascript
// JavaScript Code to set font size in cell B1 using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9, length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the specified characters
var oFont = oCharacters.GetFont();

// Set the font size to 18
oFont.SetSize(18);
```