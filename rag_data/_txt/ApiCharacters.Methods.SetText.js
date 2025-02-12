# Description / Описание

**English:** This example sets the text for the specified characters.

**Russian:** Этот пример устанавливает текст для указанных символов.

```javascript
// JavaScript code to set text for specified characters
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(23, 4);
oCharacters.SetText("string");
```

```vba
' VBA code to set text for specified characters
Sub SetSpecifiedCharacters()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")

    ' Set the value of B1
    oRange.Value = "This is just a sample text."

    ' Get characters starting at position 23 with length 4
    Set oCharacters = oRange.Characters(Start:=23, Length:=4)

    ' Set the text for the specified characters
    oCharacters.Text = "stri" ' 'stri' is 4 characters long
End Sub
```