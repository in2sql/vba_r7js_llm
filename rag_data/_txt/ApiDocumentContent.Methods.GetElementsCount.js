# Description / Описание

**English:**
This code demonstrates how to add a shape to the active worksheet, insert text with line breaks inside the shape, and retrieve the number of elements within the shape content.

**Russian:**
Этот код демонстрирует, как добавить фигуру на активный лист, вставить текст с разрывами строк внутрь фигуры и получить количество элементов внутри содержимого фигуры.

```vba
' VBA code equivalent to the OnlyOffice JS example

Dim oWorksheet As Object
Dim oFill As Object
Dim oStroke As Object
Dim oShape As Object
Dim oDocContent As Object
Dim oParagraph As Object

' Get the active sheet
Set oWorksheet = Api.GetActiveSheet()

' Create a solid fill with RGB color
Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))

' Create a stroke with no fill
Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())

' Add a shape to the worksheet
Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)

' Get the document content of the shape
Set oDocContent = oShape.GetContent()

' Get the first paragraph element
Set oParagraph = oDocContent.GetElement(0)

' Add text to the paragraph
oParagraph.AddText "We got the first paragraph inside the shape."
oParagraph.AddLineBreak
oParagraph.AddText "Number of elements inside the shape: " & oDocContent.GetElementsCount()
oParagraph.AddLineBreak
oParagraph.AddText "Line breaks are NOT counted into the number of elements."
```

```javascript
// JavaScript code using OnlyOffice API

// Get the active sheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the document content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("We got the first paragraph inside the shape.");
oParagraph.AddLineBreak();
oParagraph.AddText("Number of elements inside the shape: " + oDocContent.GetElementsCount());
oParagraph.AddLineBreak();
oParagraph.AddText("Line breaks are NOT counted into the number of elements.");
```