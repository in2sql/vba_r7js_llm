# Description / Описание

**English:** This script sets the value of cell B1 to a sample text, makes a portion of that text bold, and then displays the bold property status in cell B3.

**Русский:** Этот скрипт устанавливает значение ячейки B1 на пример текста, делает часть этого текста жирным, а затем отображает статус свойства жирного шрифта в ячейке B3.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set value to B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 to 4 characters long
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Set the font to bold
oFont.SetBold(true);

// Retrieve the bold property
var bBold = oFont.GetBold();

// Set the bold property status in cell B3
oWorksheet.GetRange("B3").SetValue("Bold property: " + bBold);
```

```vba
' VBA Code Equivalent

Sub SetBoldProperty()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    
    ' Set value to B1
    oRange.Value = "This is just a sample text."
    
    ' Get characters from position 9 to 4 characters long
    Dim oCharacters As Range
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    ' Set the font to bold
    oCharacters.Font.Bold = True
    
    ' Retrieve the bold property
    Dim bBold As Boolean
    bBold = oCharacters.Font.Bold
    
    ' Set the bold property status in cell B3
    oWorksheet.Range("B3").Value = "Bold property: " & bBold
End Sub
```