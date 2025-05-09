## Description / Описание

**English:** This code sets the value of cell B1 to a sample text and then replaces a substring within that text.

**Russian:** Этот код устанавливает значение ячейки B1 на пример текста, а затем заменяет подстроку в этом тексте.

```vba
' VBA code to set value and modify characters in a range

Sub ModifyCellText()
    Dim ws As Worksheet
    Dim rng As Range
    Dim originalText As String
    Dim newText As String

    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    ' Get the range B1
    Set rng = ws.Range("B1")
    ' Set the value of B1
    rng.Value = "This is just a sample text."
    
    ' Get the original text
    originalText = rng.Value
    ' Replace characters from position 23, length 4 with "string"
    newText = Left(originalText, 22) & "string" & Mid(originalText, 27)
    ' Update the range with the new text
    rng.Value = newText
End Sub
```

```javascript
// This example sets the text for the specified characters.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(23, 4);
oCharacters.SetText("string"); 
```