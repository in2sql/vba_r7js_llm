**Description**

*English:* This example replaces each paragraph (or text in a cell) in the selection with the corresponding text from an array of strings.

*Russian:* Этот пример заменяет каждый абзац (или текст в ячейке) в выделении соответствующим текстом из массива строк.

```vba
' VBA Code to replace cell texts
Sub ReplaceTextSmartExample()
    ' This example replaces each cell in the range A1:A2 with the corresponding text from an array of strings.

    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    oWorksheet.Range("A1").Value = "1" ' Set A1 to "1"
    oWorksheet.Range("A2").Value = "2" ' Set A2 to "2"
    
    Dim oRange As Range
    Set oRange = oWorksheet.Range("A1:A2") ' Get range A1:A2
    oRange.Select ' Select range

    ' Define array of replacement texts
    Dim texts As Variant
    texts = Array("Cell 1", "Cell 2")
    
    ' Replace each cell's value with corresponding text from array
    oRange.Cells(1, 1).Value = texts(0)
    oRange.Cells(2, 1).Value = texts(1)
End Sub
```

```javascript
// OnlyOffice JS code to replace cell texts
// This example replaces each paragraph (or text in a cell) in the selection with the corresponding text from an array of strings.

var oWorksheet = Api.GetActiveSheet(); // Get active sheet
oWorksheet.GetRange("A1").SetValue("1"); // Set A1 to "1"
oWorksheet.GetRange("A2").SetValue("2"); // Set A2 to "2"

var oRange = oWorksheet.GetRange("A1:A2"); // Get range A1:A2
oRange.Select(); // Select range

Api.ReplaceTextSmart(["Cell 1", "Cell 2"]); // Replace text smartly
```