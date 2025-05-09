# Description / Описание

**English:** This code demonstrates how to set and retrieve the superscript property of a specific portion of text in a worksheet.

**Russian:** Этот код демонстрирует, как установить и получить свойство надстрочного знака для определенной части текста в рабочем листе.

```javascript
// JavaScript code for OnlyOffice API
// This example shows how to get the superscript property of the specified font.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(9, 4);
var oFont = oCharacters.GetFont();
oFont.SetSuperscript(true);
var bSuperscript = oFont.GetSuperscript();
oWorksheet.GetRange("B3").SetValue("Superscript property: " + bSuperscript);
```

```vba
' VBA code for Excel
' This macro demonstrates how to set and retrieve the superscript property of a specific portion of text in a worksheet.

Sub SetSuperscriptProperty()
    Dim ws As Worksheet
    Dim rng As Range
    Dim characters As Characters
    Dim bSuperscript As Boolean
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Get range B1 and set its value
    Set rng = ws.Range("B1")
    rng.Value = "This is just a sample text."
    
    ' Get characters 9 to 12 (4 characters) from B1
    Set characters = rng.Characters(Start:=9, Length:=4)
    
    ' Set superscript property to True
    characters.Font.Superscript = True
    
    ' Get the superscript property
    bSuperscript = characters.Font.Superscript
    
    ' Set the value in B3
    ws.Range("B3").Value = "Superscript property: " & bSuperscript
End Sub
```