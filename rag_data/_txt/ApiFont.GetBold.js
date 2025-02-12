## Description / Описание

**English:** This code sets the value of cell B1 to a sample text, makes specific characters bold, and writes the bold property value to cell B3.

**Russian:** Этот код устанавливает значение ячейки B1 примерным текстом, делает определенные символы жирными и записывает значение свойства жирности в ячейку B3.

### Excel VBA Code

```vba
' This VBA macro sets the value of cell B1, makes specific characters bold, and displays the bold property in cell B3.

Sub SetBoldProperty()
    Dim ws As Worksheet
    Dim rng As Range
    Dim charStart As Integer
    Dim charLength As Integer
    Dim bBold As Boolean

    ' Set reference to the active sheet
    Set ws = ThisWorkbook.ActiveSheet

    ' Set value to cell B1
    Set rng = ws.Range("B1")
    rng.Value = "This is just a sample text."

    ' Define the start and length of characters to format
    charStart = 9
    charLength = 4

    ' Make specified characters bold
    With rng.Characters(Start:=charStart, Length:=charLength).Font
        .Bold = True
    End With

    ' Get the bold property
    bBold = rng.Characters(Start:=charStart, Length:=charLength).Font.Bold

    ' Set the value in B3
    ws.Range("B3").Value = "Bold property: " & bBold
End Sub
```

### OnlyOffice JavaScript Code

```javascript
// This example shows how to get the bold property of the specified font.
// Этот пример показывает, как получить свойство жирности указанного шрифта.

var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
var oRange = oWorksheet.GetRange("B1"); // Get range B1
oRange.SetValue("This is just a sample text."); // Set value to B1
var oCharacters = oRange.GetCharacters(9, 4); // Get characters starting at position 9, length 4
var oFont = oCharacters.GetFont(); // Get the font of these characters
oFont.SetBold(true); // Set bold to true
var bBold = oFont.GetBold(); // Get the bold property
oWorksheet.GetRange("B3").SetValue("Bold property: " + bBold); // Set value to B3
```