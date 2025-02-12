**Description / Описание:**

*This example sets the bold property to the specified font.*

*Этот пример устанавливает свойство жирного шрифта для указанного текста.*

---

**Excel VBA Code:**

```vba
' This example sets the bold property to the specified font.

Sub SetBoldFont()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Get the range B1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the value of B1
    oRange.Value = "This is just a sample text."
    
    ' Set characters from position 9, length 4 to bold
    oRange.Characters(Start:=9, Length:=4).Font.Bold = True
End Sub
```

---

**OnlyOffice JS Code:**

```javascript
// This example sets the bold property to the specified font.

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of B1
oRange.SetValue("This is just a sample text.");

// Get characters from position 9, length 4
var oCharacters = oRange.GetCharacters(9, 4);

// Get the font of the characters
var oFont = oCharacters.GetFont();

// Set the Bold property to true
oFont.SetBold(true);
```