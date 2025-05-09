**Description / Описание:**

This code demonstrates how to set a value in a cell, extract a substring from the cell's text, and write the extracted text to another cell.  
Этот код демонстрирует, как установить значение в ячейке, извлечь подстроку из текста ячейки и записать извлеченный текст в другую ячейку.

---

```vba
' VBA Code
Sub GetTextFromRange()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim sText As String

    ' Get the active worksheet
    Set oWorksheet = ActiveSheet

    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")

    ' Set value in B1
    oRange.Value = "This is just a sample text."

    ' Extract a substring starting at character 23 with length 4
    sText = Mid(oRange.Value, 23, 4)

    ' Set the extracted text in B3
    oWorksheet.Range("B3").Value = "Text: " & sText
End Sub
```

```javascript
// OnlyOffice JS Code
// This example shows how to get the text of the specified range of characters.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oRange = oWorksheet.GetRange("B1"); // Get the range B1
oRange.SetValue("This is just a sample text."); // Set value in B1
var oCharacters = oRange.GetCharacters(23, 4); // Get characters starting at 23, length 4
var sText = oCharacters.GetText(); // Get the extracted text
oWorksheet.GetRange("B3").SetValue("Text: " + sText); // Set the extracted text in B3
```