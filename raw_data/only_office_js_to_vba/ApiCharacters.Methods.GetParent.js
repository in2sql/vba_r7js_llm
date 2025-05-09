**Description:**

*English:* This code retrieves the active worksheet, sets a value in cell B1, selects specific characters within that cell, retrieves their parent object, and applies a thick bottom border with a specified color to that parent object.

*Russian:* Этот код получает активный рабочий лист, устанавливает значение в ячейке B1, выбирает определенные символы внутри этой ячейки, получает объект-родитель этих символов и применяет толстую нижнюю границу с заданным цветом к этому объекту.

```vba
' VBA code equivalent to the OnlyOffice JS example
Sub SetBordersExample()
    Dim oWorksheet As Worksheet
    Dim oRange As Range
    
    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Get the range B1
    Set oRange = oWorksheet.Range("B1")
    
    ' Set the value in cell B1
    oRange.Value = "This is just a sample text."
    
    ' Apply thick bottom border with specified color
    With oRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(255, 111, 61)
    End With
End Sub
```

```javascript
// JavaScript code equivalent to the OnlyOffice JS example
function setBordersExample() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();
    
    // Get the range B1
    var oRange = oWorksheet.GetRange("B1");
    
    // Set the value in cell B1
    oRange.SetValue("This is just a sample text.");
    
    // Get specific characters from position 23, length 4
    var oCharacters = oRange.GetCharacters(23, 4);
    
    // Get the parent object of the characters
    var oParent = oCharacters.GetParent();
    
    // Set thick bottom border with specified color
    oParent.SetBorders("Bottom", "Thick", Api.CreateColorFromRGB(255, 111, 61));
}
```