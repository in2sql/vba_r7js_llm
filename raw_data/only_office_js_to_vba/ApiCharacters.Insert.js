# Insert a String Replacing the Specified Characters
# Вставка строки с заменой указанных символов

```vba
' VBA code to insert a string by replacing specified characters.

Sub InsertString()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim startPos As Integer
    Dim length As Integer
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")
    ' Set the value of B1
    oRange.Value = "This is just a sample text."
    
    startPos = 23
    length = 4
    ' Insert "string" starting at position 23, replacing 4 characters
    oRange.Characters(Start:=startPos, Length:=length).Insert "string"
End Sub
```

```javascript
// This example inserts a string replacing the specified characters.
// Этот пример вставляет строку, заменяя указанные символы.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set the value of B1
var oCharacters = oRange.GetCharacters(23, 4); // Get characters starting at position 23 with length 4
// Insert "string" starting at position 23, replacing 4 characters
oCharacters.Insert("string"); 
```