# Description
**English:** This code demonstrates how to set a specific underline style to a portion of text in cell B1 and retrieve its underline property.

**Russian:** Этот код демонстрирует, как установить определённый стиль подчёркивания для части текста в ячейке B1 и получить его свойство подчёркивания.

```javascript
// JavaScript OnlyOffice API code
// This code sets an underline style to a portion of text in cell B1 and retrieves the underline property.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get range B1
oRange.SetValue("This is just a sample text."); // Set value in B1
var oCharacters = oRange.GetCharacters(9, 4); // Get characters from position 9, length 4
var oFont = oCharacters.GetFont(); // Get font of selected characters
oFont.SetUnderline("xlUnderlineStyleSingle"); // Set underline style to single
var sUnderline = oFont.GetUnderline(); // Get underline style
oWorksheet.GetRange("B3").SetValue("Underline property: " + sUnderline); // Set value in B3 with underline property
```

```vba
' VBA code
' This code sets an underline style to a portion of text in cell B1 and retrieves the underline property.

Sub SetUnderlineExample()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters
    Dim sUnderline As String
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get range B1
    Set oRange = oWorksheet.Range("B1")
    
    ' Set value in B1
    oRange.Value = "This is just a sample text."
    
    ' Get characters from position 9, length 4
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    ' Set underline style to single
    oCharacters.Font.Underline = xlUnderlineStyleSingle
    
    ' Get underline style
    sUnderline = oCharacters.Font.Underline
    
    ' Set value in B3 with underline property
    oWorksheet.Range("B3").Value = "Underline property: " & sUnderline
End Sub
```