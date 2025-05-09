**Description / Описание**

This code sets an underline style to a specific portion of text in cell B1 of the active worksheet.

Этот код устанавливает стиль подчеркивания для определенной части текста в ячейке B1 активного рабочего листа.

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9 to 12
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the selected characters
var oFont = oCharacters.GetFont();

// Set the underline style to single underline
oFont.SetUnderline("xlUnderlineStyleSingle");
```

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Get the range B1
Dim oRange As Range
Set oRange = oWorksheet.Range("B1")

' Set the value of B1
oRange.Value = "This is just a sample text."

' Get characters from position 9 to 12
Dim oCharacters As Characters
Set oCharacters = oRange.Characters(Start:=9, Length:=4)

' Set the underline style to single underline
oCharacters.Font.Underline = xlUnderlineStyleSingle
```