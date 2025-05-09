**Description / Описание**

English: This code demonstrates how to find the intersection of two ranges in a worksheet and set the fill color of the intersecting range.

Russian: Этот код демонстрирует, как найти пересечение двух диапазонов в листе и установить цвет заливки пересекающегося диапазона.

```vba
' VBA code to find the intersection of two ranges and set the fill color

Sub SetIntersectionFillColor()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define the first range
    Dim oRange1 As Range
    Set oRange1 = oWorksheet.Range("A1:C5")
    
    ' Define the second range
    Dim oRange2 As Range
    Set oRange2 = oWorksheet.Range("B2:B4")
    
    ' Find the intersection of the two ranges
    Dim oIntersection As Range
    Set oIntersection = Application.Intersect(oRange1, oRange2)
    
    ' Check if the intersection is not nothing
    If Not oIntersection Is Nothing Then
        ' Set the fill color using RGB values
        oIntersection.Interior.Color = RGB(255, 213, 191)
    End If
End Sub
```

```javascript
// JavaScript code to find the intersection of two ranges and set the fill color

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Define the first range
var oRange1 = oWorksheet.GetRange("A1:C5");

// Define the second range
var oRange2 = oWorksheet.GetRange("B2:B4");

// Find the intersection of the two ranges
var oRange = Api.Intersect(oRange1, oRange2);

// Set the fill color using RGB values
oRange.SetFillColor(Api.CreateColorFromRGB(255, 213, 191)); 
```