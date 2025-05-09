## Description
This code demonstrates how to set and retrieve the superscript property of specific characters in a cell in Excel using OnlyOffice API and its equivalent in Excel VBA.

Этот код демонстрирует, как установить и получить свойство верхнего индекса для определенных символов в ячейке Excel, используя OnlyOffice API и его эквивалент в Excel VBA.

```javascript
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
' This example shows how to get the superscript property of the specified font.
Sub SetSuperscript()
    Dim ws As Worksheet
    Dim rng As Range
    Dim chars As Characters
    Dim isSuperscript As Boolean
    
    Set ws = ThisWorkbook.ActiveSheet
    Set rng = ws.Range("B1")
    
    rng.Value = "This is just a sample text."
    
    ' Characters in VBA are 1-based
    Set chars = rng.Characters(Start:=9, Length:=4)
    chars.Font.Superscript = True
    
    isSuperscript = chars.Font.Superscript
    
    ws.Range("B3").Value = "Superscript property: " & isSuperscript
End Sub
```