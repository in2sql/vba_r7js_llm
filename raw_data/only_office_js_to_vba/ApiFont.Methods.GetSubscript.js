### Description
**English:** This code demonstrates how to set and retrieve the subscript property of a specified font in a cell.  
**Russian:** Этот код демонстрирует, как установить и получить свойство подстрочного шрифта в указанной ячейке.

```javascript
// JavaScript OnlyOffice API Code

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set value in B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9, length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font object from characters
var oFont = oCharacters.GetFont();

// Set subscript to true
oFont.SetSubscript(true);

// Get the subscript property
var bSubscript = oFont.GetSubscript();

// Set value in B3 with subscript property
oWorksheet.GetRange("B3").SetValue("Subscript property: " + bSubscript); 
```

```vba
' VBA Code Equivalent

' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Get the range B1
Dim oRange As Range
Set oRange = oWorksheet.Range("B1")

' Set value in B1
oRange.Value = "This is just a sample text."

' Get characters from position 9, length 4
Dim oCharacters As Characters
Set oCharacters = oRange.Characters(Start:=9, Length:=4)

' Get the font object from characters
Dim oFont As Font
Set oFont = oCharacters.Font

' Set subscript to true
oFont.Subscript = True

' Get the subscript property
Dim bSubscript As Boolean
bSubscript = oFont.Subscript

' Set value in B3 with subscript property
oWorksheet.Range("B3").Value = "Subscript property: " & bSubscript
```