**Description / Описание:**

English: This script sets the value of cell B1 to a sample text, retrieves characters starting at position 9 with a length of 4 from the cell, and deletes those characters.

Russian: Этот скрипт устанавливает значение ячейки B1 на пример текста, извлекает символы, начиная с позиции 9 длиной 4 символа, из ячейки и удаляет эти символы.

---

**VBA Code:**

```vba
' VBA Script to set a value in cell B1, retrieve specific characters, and delete them

Sub DeleteCharactersInCell()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cellValue As String
    Dim startPos As Integer
    Dim lengthToDelete As Integer
    
    ' Set the worksheet to the active sheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Set the range to cell B1
    Set rng = ws.Range("B1")
    
    ' Set the value of cell B1
    rng.Value = "This is just a sample text."
    
    ' Get the current value of the cell
    cellValue = rng.Value
    
    ' Define the starting position and number of characters to delete
    startPos = 9
    lengthToDelete = 4
    
    ' Delete the specified characters from the cell value
    If Len(cellValue) >= (startPos + lengthToDelete - 1) Then
        rng.Value = Left(cellValue, startPos - 1) & Mid(cellValue, startPos + lengthToDelete)
    Else
        MsgBox "The specified range exceeds the length of the cell content."
    End If
End Sub
```

---

**OnlyOffice JS Code:**

```javascript
// JavaScript to set a value in cell B1, retrieve specific characters, and delete them

function deleteCharactersInCell() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Get the range for cell B1
    var oRange = oWorksheet.GetRange("B1");
    
    // Set the value of cell B1
    oRange.SetValue("This is just a sample text.");
    
    // Get characters starting at position 9 with a length of 4
    var oCharacters = oRange.GetCharacters(9, 4);
    
    // Delete the retrieved characters
    oCharacters.Delete();
}
```