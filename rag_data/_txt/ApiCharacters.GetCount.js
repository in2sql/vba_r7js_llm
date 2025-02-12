**Description**

*English:* This code retrieves the active worksheet, sets the value of cell B1 with a sample text, extracts a subset of characters from the text in B1, counts the number of characters extracted, and then sets the value of cell B3 to display the number of characters.

*Russian:* Этот код получает активный рабочий лист, устанавливает значение ячейки B1 с примерным текстом, извлекает подмножество символов из текста в B1, считает количество извлеченных символов и затем устанавливает значение ячейки B3 для отображения количества символов.

---

**Excel VBA Code**

```vba
' This VBA macro performs similar operations as the OnlyOffice JS example.

Sub CountCharacters()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim strText As String
    Dim strSubset As String
    Dim nCount As Integer
    
    ' Get the active worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value of cell B1
    Set oRange = oWorksheet.Range("B1")
    oRange.Value = "This is just a sample text."
    
    ' Get the text from B1
    strText = oRange.Value
    
    ' Extract characters from position 23 to 26
    If Len(strText) >= 26 Then
        strSubset = Mid(strText, 23, 4)
    Else
        strSubset = Mid(strText, 23, Len(strText) - 22)
    End If
    
    ' Count the number of characters in the subset
    nCount = Len(strSubset)
    
    ' Set value of cell B3 with the count
    oWorksheet.Range("B3").Value = "Number of characters: " & nCount
End Sub
```

---

**OnlyOffice JS Code**

```javascript
// This OnlyOffice JS code retrieves the active worksheet, sets cell B1 with sample text,
// extracts a subset of characters, counts them, and sets cell B3 with the character count.

function countCharacters() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value of cell B1
    var oRange = oWorksheet.GetRange("B1");
    oRange.SetValue("This is just a sample text.");
    
    // Get a subset of characters from position 23, length 4
    var oCharacters = oRange.GetCharacters(23, 4);
    
    // Count the number of characters
    var nCount = oCharacters.GetCount();
    
    // Set value of cell B3 with the count
    oWorksheet.GetRange("B3").SetValue("Number of characters: " + nCount);
}

// Execute the function
countCharacters();
```