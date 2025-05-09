### Description / Описание

**English:** This code sets the value of cell B1 and applies subscript formatting to characters 9 to 12.

**Русский:** Этот код устанавливает значение ячейки B1 и применяет форматирование нижнего индекса к символам с 9 по 12.

```javascript
// JavaScript Code for OnlyOffice API
// This example sets the subscript property to the specified font.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(9, 4);
var oFont = oCharacters.GetFont();
oFont.SetSubscript(true); 
```

```vba
' VBA Code Equivalent
' This example sets the subscript property to the specified font.

Sub SetSubscript()
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    
    oRange.Value = "This is just a sample text."
    
    ' Apply subscript to characters 9 to 12
    oRange.Characters(Start:=9, Length:=4).Font.Subscript = True
End Sub
```