# Insert a string by replacing specified characters
# Вставляет строку, заменяя указанные символы

```javascript
// JavaScript code using OnlyOffice API
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(23, 4);
oCharacters.Insert("string"); 
```

```vba
' VBA code equivalent
Sub InsertStringInCell()
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Dim rng As Range
    Set rng = ws.Range("B1")
    
    ' Set value to cell B1
    rng.Value = "This is just a sample text."
    
    ' Insert "string" starting at the 23rd character for 4 characters
    Dim originalText As String
    originalText = rng.Value
    rng.Value = Left(originalText, 22) & "string" & Mid(originalText, 27)
End Sub
```