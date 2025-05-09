# Description / Описание

**English:** This script retrieves the active worksheet, creates a color from RGB values, sets a cell's value and font color, gets the color's class type, and inserts the class type into another cell.

**Russian:** Этот скрипт получает активный лист, создает цвет из RGB-значений, устанавливает значение и цвет шрифта ячейки, получает тип класса цвета и вставляет тип класса в другую ячейку.

```vba
' VBA Code to perform the equivalent operations

Sub SetCellColorAndClassType()
    ' Retrieve the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Create a color from RGB values
    Dim oColor As Long
    oColor = RGB(255, 111, 61)
    
    ' Set the value of cell A2
    oWorksheet.Range("A2").Value = "Text with color"
    
    ' Set the font color of cell A2
    oWorksheet.Range("A2").Font.Color = oColor
    
    ' Get the color class type (In VBA, we can represent it as a string)
    Dim sColorClassType As String
    sColorClassType = "RGB(" & Red(oColor) & ", " & Green(oColor) & ", " & Blue(oColor) & ")"
    
    ' Set the value of cell A4 with the color class type
    oWorksheet.Range("A4").Value = "Class type = " & sColorClassType
End Sub
```

```javascript
// JavaScript Code using OnlyOffice API to perform the equivalent operations

// Retrieve the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a color from RGB values
var oColor = Api.CreateColorFromRGB(255, 111, 61);

// Set the value of cell A2
oWorksheet.GetRange("A2").SetValue("Text with color");

// Set the font color of cell A2
oWorksheet.GetRange("A2").SetFontColor(oColor);

// Get the color class type
var sColorClassType = oColor.GetClassType();

// Set the value of cell A4 with the color class type
oWorksheet.GetRange("A4").SetValue("Class type = " + sColorClassType);
```