# Description / Описание

This code demonstrates how to set a value in cell B1, retrieve specific characters from that cell, obtain the parent range of those characters, and apply a thick bottom border with a specific color.

Этот код демонстрирует, как установить значение в ячейку B1, получить определенные символы из этой ячейки, получить родительский диапазон этих символов и применить толстую нижнюю границу с определенным цветом.

---

## Excel VBA

```vba
' Get the active worksheet
Dim oWorksheet As Worksheet
Set oWorksheet = ThisWorkbook.ActiveSheet

' Get the range B1
Dim oRange As Range
Set oRange = oWorksheet.Range("B1")

' Set the value of B1
oRange.Value = "This is just a sample text."

' Note: VBA does not support getting characters and their parents directly.
' Instead, apply border to the entire range or specific parts as needed.

' Set the bottom border of the range with Thick line and specified color
With oRange.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThick
    .Color = RGB(255, 111, 61)
End With
```

## OnlyOffice JS

```javascript
// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Get the range B1
var oRange = oWorksheet.GetRange("B1");

// Set the value of B1
oRange.SetValue("This is just a sample text.");

// Get characters starting from index 23 with length 4
var oCharacters = oRange.GetCharacters(23, 4);

// Get the parent object of the characters
var oParent = oCharacters.GetParent();

// Set the bottom border with Thick style and specified color
oParent.SetBorders("Bottom", "Thick", Api.CreateColorFromRGB(255, 111, 61));
```