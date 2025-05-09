**Description / Описание**

This code retrieves the active worksheet, creates a shape with specified fill and stroke properties, obtains the class type of the shape's content, aligns the paragraph to the left, and adds text indicating the class type.

Этот код получает активный лист, создает фигуру с указанными свойствами заливки и обводки, получает тип класса содержимого фигуры, выравнивает абзац по левому краю и добавляет текст, указывающий тип класса.

```javascript
// JavaScript Code using OnlyOffice API

// This example gets a class type and inserts it into the document.
var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var sClassType = oDocContent.GetClassType(); // Get the class type of the content
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph element
oParagraph.SetJc("left"); // Set paragraph alignment to left
oParagraph.AddText("Class Type = " + sClassType); // Add text with the class type
```

```vba
' VBA Code Equivalent

' This macro retrieves the active worksheet, creates a shape with specific fill and stroke,
' obtains the class type of the shape's content, aligns the paragraph to the left,
' and adds text indicating the class type.

Sub AddShapeWithClassType()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet ' Get the active worksheet
    
    ' Create a color with RGB(255, 111, 61)
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a rectangle shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStorageData, 200, 60, 200, 100) ' Adjust size as needed
    
    ' Set the fill color
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = fillColor
    End With
    
    ' Set the stroke to no line
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Assuming we have a custom property or method to get class type
    ' VBA does not have a direct equivalent to GetClassType, so this is a placeholder
    Dim classType As String
    classType = "CustomClassType" ' Replace with actual method to get class type if available
    
    ' Add text to the shape
    With shp.TextFrame2
        .HorizontalAnchor = msoAnchorLeft ' Align text to left
        .TextRange.Text = "Class Type = " & classType
    End With
End Sub
```