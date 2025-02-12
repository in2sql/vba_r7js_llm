# Delete ApiCharacters Object / Удаление объекта ApiCharacters

**English:** This code deletes specific characters in cell B1.  
**Russian:** Этот код удаляет конкретные символы в ячейке B1.

```vba
' This VBA code deletes specific characters in cell B1
Sub DeleteApiCharacters()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cellText As String
    Dim newText As String
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Set rng = ws.Range("B1")
    
    ' Set the value of B1
    rng.Value = "This is just a sample text."
    
    ' Get the current text from B1
    cellText = rng.Value
    
    ' Delete 4 characters starting from the 9th character
    newText = Left(cellText, 8) & Mid(cellText, 13)
    
    ' Update B1 with the new text
    rng.Value = newText
End Sub
```

```javascript
// This example deletes the ApiCharacters object.
// Этот пример удаляет объект ApiCharacters.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set the value of B1
var oCharacters = oRange.GetCharacters(9, 4); // Get characters starting at position 9, length 4
oCharacters.Delete(); // Delete the specified characters
```