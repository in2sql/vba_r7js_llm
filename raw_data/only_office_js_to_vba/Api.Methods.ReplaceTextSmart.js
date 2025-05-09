**Description / Описание**

*English:* This example replaces each paragraph (or text in a cell) in the selected range with the corresponding text from an array of strings.

*Russian:* Этот пример заменяет каждый абзац (или текст в ячейке) в выбранном диапазоне соответствующим текстом из массива строк.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Set value "1" in cell A1
oWorksheet.GetRange("A1").SetValue("1");

// Set value "2" in cell A2
oWorksheet.GetRange("A2").SetValue("2");

// Get the range A1:A2
var oRange = oWorksheet.GetRange("A1:A2");

// Select the range
oRange.Select();

// Replace text in the selected range with the provided array
Api.ReplaceTextSmart(["Cell 1", "Cell 2"]);
```

```vba
' VBA Code Equivalent

Sub ReplaceTextInRange()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Set value "1" in cell A1
    oWorksheet.Range("A1").Value = "1"
    
    ' Set value "2" in cell A2
    oWorksheet.Range("A2").Value = "2"
    
    ' Define the range A1:A2
    Dim oRange As Range
    Set oRange = oWorksheet.Range("A1:A2")
    
    ' Select the range
    oRange.Select
    
    ' Replace text in the selected range with the provided array
    ' Since VBA does not have a direct ReplaceTextSmart method, we iterate through the cells
    Dim values As Variant
    values = Array("Cell 1", "Cell 2")
    
    Dim i As Integer
    For i = 1 To oRange.Cells.Count
        oRange.Cells(i).Value = values(i - 1)
    Next i
End Sub
```