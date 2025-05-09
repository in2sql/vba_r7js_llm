// This code sets the underline style of a specific substring in cell B1.
// Этот код устанавливает стиль подчеркивания для определенной подстроки в ячейке B1.

```vba
' VBA code equivalent to set underline style for specific characters in cell B1

Sub SetUnderlineStyle()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters
    Dim oFont As Font
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the value of B1
    oRange.Value = "This is just a sample text."
    
    ' Get characters from position 9 to 12 (4 characters)
    Set oCharacters = oRange.Characters(Start:=9, Length:=4)
    
    ' Get the font of these characters
    Set oFont = oCharacters.Font
    
    ' Set the underline style to single
    oFont.Underline = xlUnderlineStyleSingle
End Sub
```

```javascript
// This example sets an underline of the type specified in the request to the font.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set the value of B1
var oCharacters = oRange.GetCharacters(9, 4); // Get characters from position 9 with length 4
var oFont = oCharacters.GetFont(); // Get the font of these characters
oFont.SetUnderline("xlUnderlineStyleSingle"); // Set underline style to single
```