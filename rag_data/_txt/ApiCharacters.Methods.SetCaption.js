# Description / Описание

**English:**  
This code sets a string value in cell B1 and updates the caption of specific characters within that cell.

**Russian:**  
Этот код устанавливает строковое значение в ячейку B1 и обновляет заголовок определенных символов внутри этой ячейки.

```vba
' This example sets a string value that represents the text of the specified range of characters.
Sub SetRangeValueAndCaption()
    Dim oWorksheet As Worksheet
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    Dim oRange As Range
    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the value of the range
    oRange.Value = "This is just a sample text."
    
    Dim oCharacters As Characters
    ' Get characters from position 23 with length 4
    Set oCharacters = oRange.Characters(Start:=23, Length:=4)
    
    ' Set the caption by changing the text (VBA does not have a direct SetCaption method)
    oCharacters.Text = "string"
End Sub
```

```javascript
// This example sets a string value that represents the text of the specified range of characters.
function setRangeValueAndCaption() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Get the range B1
    var oRange = oWorksheet.GetRange("B1");
    
    // Set the value of the range
    oRange.SetValue("This is just a sample text.");
    
    // Get characters from position 23 with length 4
    var oCharacters = oRange.GetCharacters(23, 4);
    
    // Set the caption
    oCharacters.SetCaption("string");
}
```