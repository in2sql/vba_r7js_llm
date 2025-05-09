# Description / Описание

**English:** This code sets a sample text in cell B1, retrieves a specific range of characters from the text, counts the number of those characters, and then sets the count in cell B3.

**Russian:** Этот код устанавливает пример текста в ячейку B1, извлекает определенный диапазон символов из текста, подсчитывает количество этих символов и затем записывает их количество в ячейку B3.

```vba
' VBA Code
Sub CountCharacters()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    Dim sText As String
    Dim sSubset As String
    Dim nCount As Long
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Set value in cell B1
    Set oRange = oWorksheet.Range("B1")
    oRange.Value = "This is just a sample text."
    
    ' Get the text from B1
    sText = oRange.Value
    
    ' Extract characters from position 23 to 26 (4 characters)
    If Len(sText) >= 26 Then
        sSubset = Mid(sText, 23, 4)
        nCount = Len(sSubset)
    ElseIf Len(sText) >= 23 Then
        sSubset = Mid(sText, 23, Len(sText) - 22)
        nCount = Len(sSubset)
    Else
        sSubset = ""
        nCount = 0
    End If
    
    ' Set the count in cell B3
    oWorksheet.Range("B3").Value = "Number of characters: " & nCount
End Sub
```

```javascript
// OnlyOffice JS Code
// This example sets a sample text in cell B1, retrieves a specific range of characters, counts them, and writes the count to cell B3.

Api.OnDocumentReady(function(){
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Set value in cell B1
    var oRange = oWorksheet.GetRange("B1");
    oRange.SetValue("This is just a sample text.");
    
    // Get characters from position 23 to 26 (4 characters)
    var oCharacters = oRange.GetCharacters(23, 4);
    var nCount = oCharacters.GetCount();
    
    // Set the count in cell B3
    oWorksheet.GetRange("B3").SetValue("Number of characters: " + nCount);
});
```