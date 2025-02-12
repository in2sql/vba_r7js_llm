**English:** This script sets the value of cell B1 to "This is just a sample text." and makes the characters from the 9th to the 12th position bold.

**Русский:** Этот скрипт устанавливает значение ячейки B1 на "This is just a sample text." и делает жирными символы с 9 по 12 позицию.

```vba
' VBA code to set cell B1 value and make specific characters bold
Sub FormatCellText()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cellText As String
    
    ' Get the active worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Set rng = ws.Range("B1")
    
    ' Set the value of B1
    cellText = "This is just a sample text."
    rng.Value = cellText
    
    ' Make characters from position 9 to 12 bold
    With rng.Characters(Start:=9, Length:=4).Font
        .Bold = True
    End With
End Sub
```

```javascript
// JavaScript code to set cell B1 value and make specific characters bold using OnlyOffice API
function formatCellText() {
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
    
    // Set the selected characters to bold
    oFont.SetBold(true);
}
```