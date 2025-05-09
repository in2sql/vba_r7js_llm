# Description / Описание

This example shows how to get the text of the specified range of characters.
Этот пример показывает, как получить текст указанного диапазона символов.

## Excel VBA Code

```vba
' This example shows how to get the text of the specified range of characters.

Sub GetCharactersText()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get range B1 and set its value
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    oRange.Value = "This is just a sample text."
    
    ' Get characters from position 23, length 4
    Dim oCharacters As Characters
    Set oCharacters = oRange.Characters(Start:=23, Length:=4)
    
    ' Get text from characters
    Dim sText As String
    sText = oCharacters.Text
    
    ' Set value in B3
    oWorksheet.Range("B3").Value = "Text: " & sText
End Sub
```

## OnlyOffice JavaScript Code

```javascript
// This example shows how to get the text of the specified range of characters.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set value in B1
var oCharacters = oRange.GetCharacters(23, 4); // Get characters from position 23, length 4
var sText = oCharacters.GetText(); // Get text from characters
oWorksheet.GetRange("B3").SetValue("Text: " + sText); // Set value in B3
```