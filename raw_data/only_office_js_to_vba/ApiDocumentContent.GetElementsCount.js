**Description / Описание:**  
This code creates a shape in the active worksheet, adds a paragraph with text and line breaks, and displays the number of elements inside the shape.  
Этот код создает фигуру на активном листе, добавляет абзац с текстом и разрывами строк, а также отображает количество элементов внутри фигуры.

```javascript
// JavaScript Code using OnlyOffice API

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape(
    "flowChartOnlineStorage", 
    200 * 36000, 
    60 * 36000, 
    oFill, 
    oStroke, 
    0, 
    2 * 36000, 
    0, 
    3 * 36000
);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Add text to the paragraph
oParagraph.AddText("We got the first paragraph inside the shape.");

// Add a line break
oParagraph.AddLineBreak();

// Add text displaying the number of elements inside the shape
oParagraph.AddText("Number of elements inside the shape: " + oDocContent.GetElementsCount());

// Add another line break
oParagraph.AddLineBreak();

// Add explanatory text
oParagraph.AddText("Line breaks are NOT counted into the number of elements.");
```

```vba
' VBA Code Equivalent

Sub CreateShapeAndAddContent()
    ' Get the active worksheet
    Dim oWorksheet As Object
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Dim oFill As Object
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Dim oStroke As Object
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet with specified parameters
    Dim oShape As Object
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 200 * 36000, 60 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Dim oDocContent As Object
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph element
    Dim oParagraph As Object
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Add text to the paragraph
    Call oParagraph.AddText("We got the first paragraph inside the shape.")
    
    ' Add a line break
    Call oParagraph.AddLineBreak()
    
    ' Add text displaying the number of elements inside the shape
    Call oParagraph.AddText("Number of elements inside the shape: " & oDocContent.GetElementsCount())
    
    ' Add another line break
    Call oParagraph.AddLineBreak()
    
    ' Add explanatory text
    Call oParagraph.AddText("Line breaks are NOT counted into the number of elements.")
End Sub
```