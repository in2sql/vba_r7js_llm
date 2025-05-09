### Description / Описание
This script adds a shape to the active worksheet, sets its fill and stroke properties, aligns text to the left within the shape, and inserts a line break between two text elements.
Этот скрипт добавляет фигуру на активный лист, устанавливает свойства заливки и обводки, выравнивает текст по левому краю внутри фигуры и вставляет разрыв строки между двумя текстовыми элементами.

```vba
' VBA Code to add a shape with specific properties and text formatting

Sub AddShapeWithText()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ActiveSheet
    
    ' Define fill color (RGB: 255, 111, 61)
    Dim oFill As Object
    Set oFill = CreateSolidFill(255, 111, 61)
    
    ' Define stroke (no fill)
    Dim oStroke As Object
    Set oStroke = CreateStroke(0, CreateNoFill())
    
    ' Add shape to the worksheet
    Dim oShape As Object
    Set oShape = oWorksheet.Shapes.AddShape( _
        Type:=msoShapeFlowchartDatabase, _
        Left:=120, _
        Top:=70, _
        Width:=200, _
        Height:=100)
    
    ' Set fill and stroke properties
    With oShape
        .Fill.ForeColor.RGB = RGB(255, 111, 61)
        .Line.Visible = msoFalse
    End With
    
    ' Add text to the shape
    With oShape.TextFrame2.TextRange
        .ParagraphFormat.Alignment = msoAlignLeft
        .Text = "This is a text inside the shape aligned left." & vbCrLf & "This is a text after the line break."
    End With
End Sub

' Function to create a solid fill
Function CreateSolidFill(red As Integer, green As Integer, blue As Integer) As Object
    Dim fill As Object
    Set fill = CreateObject("OnlyOffice.Fill")
    fill.Type = "Solid"
    fill.Color = RGB(red, green, blue)
    Set CreateSolidFill = fill
End Function

' Function to create a stroke with no fill
Function CreateStroke(width As Integer, fill As Object) As Object
    Dim stroke As Object
    Set stroke = CreateObject("OnlyOffice.Stroke")
    stroke.Width = width
    Set stroke.Fill = fill
    Set CreateStroke = stroke
End Function

' Function to create no fill
Function CreateNoFill() As Object
    Dim fill As Object
    Set fill = CreateObject("OnlyOffice.Fill")
    fill.Type = "None"
    Set CreateNoFill = fill
End Function
```

```javascript
// JavaScript Code to add a shape with specific properties and text formatting

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Set paragraph alignment to left
oParagraph.SetJc("left");

// Add text to the paragraph
oParagraph.AddText("This is a text inside the shape aligned left.");

// Add a line break
oParagraph.AddLineBreak();

// Add more text after the line break
oParagraph.AddText("This is a text after the line break.");
```