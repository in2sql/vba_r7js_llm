**Description / Описание**  
This code demonstrates how to set and retrieve the strikethrough property of specific characters in a cell.  
Этот код демонстрирует, как установить и получить свойство зачёркивания для определённых символов в ячейке.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 with length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Apply strikethrough to the font
oFont.SetStrikethrough(true);

// Retrieve the strikethrough property
var bStrikethrough = oFont.GetStrikethrough();

// Set the value of B3 to display the strikethrough property
oWorksheet.GetRange("B3").SetValue("Strikethrough property: " + bStrikethrough);
```

```vba
' VBA Code Equivalent

Sub SetAndGetStrikethrough()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set the value of B1
    oWorksheet.Range("B1").Value = "This is just a sample text."
    
    ' Apply strikethrough to characters 9 to 12 in B1
    With oWorksheet.Range("B1").Characters(Start:=9, Length:=4).Font
        .Strikethrough = True
        ' Retrieve the strikethrough property
        Dim bStrikethrough As Boolean
        bStrikethrough = .Strikethrough
    End With
    
    ' Set the value of B3 to display the strikethrough property
    oWorksheet.Range("B3").Value = "Strikethrough property: " & bStrikethrough
End Sub
```