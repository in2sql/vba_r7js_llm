# Description / Описание

**English:** This code sets the superscript property for a specific portion of text in cell B1 of the active worksheet.

**Russian:** Этот код устанавливает свойство надстрочного написания для определенной части текста в ячейке B1 активного листа.

```vba
' VBA code equivalent to set superscript in cell B1
Sub SetSuperscript()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Dim oRange As Range
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the cell value
    oRange.Value = "This is just a sample text."
    
    ' Set superscript for characters 9 to 12 (4 characters starting at 9)
    With oRange.Characters(Start:=9, Length:=4).Font
        .Superscript = True
    End With
End Sub
```

```javascript
// JavaScript code equivalent to set superscript in cell B1
function setSuperscript() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Get the range B1
    var oRange = oWorksheet.GetRange("B1");
    
    // Set the cell value
    oRange.SetValue("This is just a sample text.");
    
    // Get characters starting at position 9, length 4
    var oCharacters = oRange.GetCharacters(9, 4);
    
    // Get the font of the characters
    var oFont = oCharacters.GetFont();
    
    // Set superscript
    oFont.SetSuperscript(true);
}
```