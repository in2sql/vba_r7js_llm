**Description / Описание**

This code sets the value and font color of cell A2 and displays the RGB color value in cell A4.

Этот код устанавливает значение и цвет шрифта ячейки A2 и отображает RGB значение цвета в ячейке A4.

```vba
' VBA Code to set cell value and font color, then display RGB color value

Sub SetCellColorAndValue()
    Dim oWorksheet As Worksheet
    Dim rCellA2 As Range
    Dim rCellA4 As Range
    Dim colorRGB As Long

    ' Get the active worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet

    ' Define the range A2 and A4
    Set rCellA2 = oWorksheet.Range("A2")
    Set rCellA4 = oWorksheet.Range("A4")

    ' Set value in A2
    rCellA2.Value = "Text with color"

    ' Set font color in A2 using RGB
    rCellA2.Font.Color = RGB(255, 111, 61)

    ' Get the RGB color value
    colorRGB = rCellA2.Font.Color

    ' Set value in A4 with the RGB color value
    rCellA4.Value = "Cell color in RGB format: " & colorRGB
End Sub
```

```javascript
// JavaScript Code to set cell value and font color, then display RGB color value

function setCellColorAndValue() {
    // Get the active worksheet
    var oWorksheet = Api.GetActiveSheet();

    // Create a color from RGB values
    var oColor = Api.CreateColorFromRGB(255, 111, 61);

    // Set the value of cell A2
    oWorksheet.GetRange("A2").SetValue("Text with color");

    // Set the font color of cell A2
    oWorksheet.GetRange("A2").SetFontColor(oColor);

    // Get the RGB value of the color
    var nColor = oColor.GetRGB();

    // Set the value of cell A4 with the RGB color value
    oWorksheet.GetRange("A4").SetValue("Cell color in RGB format: " + nColor);
}
```