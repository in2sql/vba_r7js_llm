### Description / Описание

**English:**  
This code sets a string value in cell B1 of the active worksheet and then modifies a specific range of characters within that cell to have a new caption.

**Russian:**  
Этот код устанавливает строковое значение в ячейку B1 активного листа и затем изменяет определенный диапазон символов в этой ячейке, задавая новую подпись.

---

#### VBA Code

```vba
Sub SetValueAndModifyCharacters()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim oCharacters As Characters
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the value of the range
    oRange.Value = "This is just a sample text."
    
    ' Check if the text is long enough
    If Len(oRange.Value) >= 26 Then
        ' Get the characters from position 23, length 4
        Set oCharacters = oRange.Characters(Start:=23, Length:=4)
        
        ' Set the text of the specified characters
        oCharacters.Text = "string"
    Else
        MsgBox "The text is not long enough to modify characters."
    End If
End Sub
```

---

#### OnlyOffice JS Code

```javascript
// This example sets a string value that represents the text of the specified range of characters.
var oWorksheet = Api.GetActiveSheet();
var oRange = oWorksheet.GetRange("B1");
oRange.SetValue("This is just a sample text.");
var oCharacters = oRange.GetCharacters(23, 4);
oCharacters.SetCaption("string");
```