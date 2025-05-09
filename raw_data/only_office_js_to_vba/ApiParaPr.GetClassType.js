**Description / Описание:**

This code retrieves the active worksheet, creates a solid fill and stroke, adds a flowchart shape with specific dimensions, modifies paragraph properties, adds text to the shape, and appends a new paragraph showing the class type.

Этот код получает активный лист, создает заливку и обводку, добавляет фигуру блок-схемы с определенными размерами, изменяет свойства абзаца, добавляет текст к фигуре и добавляет новый абзац с отображением типа класса.

```javascript
// This example gets the active worksheet, creates a fill and stroke, 
// adds a flowchart shape, modifies paragraph properties, adds text, 
// and appends a new paragraph with the class type.

var oWorksheet = Api.GetActiveSheet(); // Get the active worksheet

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); 

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); 

// Add a flowchart shape to the worksheet with specified dimensions and styles
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); 

// Get the content of the shape
var oDocContent = oShape.GetContent(); 

// Get the first paragraph element from the content
var oParagraph = oDocContent.GetElement(0); 

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr(); 

// Retrieve the class type of the paragraph
var sClassType = oParaPr.GetClassType(); 

// Set the first line indent to 1440
oParaPr.SetIndFirstLine(1440); 

// Add multiple text strings to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 1 inch set to the first line. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes."); 

// Create a new paragraph
oParagraph = Api.CreateParagraph(); 

// Add text displaying the class type
oParagraph.AddText("Class Type = " + sClassType); 

// Append the new paragraph to the document content
oDocContent.Push(oParagraph); 
```

```vba
' This VBA code retrieves the active worksheet, creates a fill and stroke, 
' adds a flowchart shape, modifies paragraph properties, adds text, 
' and appends a new paragraph with the class type.

Sub AddFlowChartShape()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Define the fill color using RGB values
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add a flowchart shape to the worksheet with specified dimensions
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartOfflineStorage, _
        Left:=120, Top:=70, Width:=200, Height:=100) ' Adjust dimensions as needed
    
    ' Set the fill color of the shape
    With oShape.Fill
        .Visible = msoTrue
        .ForeColor.RGB = fillColor
        .Solid
    End With
    
    ' Remove the stroke (outline) of the shape
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Add text to the shape
    Dim sClassType As String
    sClassType = "FlowChartType" ' Example class type
    
    ' Set paragraph formatting for the shape's text
    With oShape.TextFrame2.TextRange.ParagraphFormat
        .FirstLineIndent = 1440 ' Set first line indent to 1 inch (1440 points)
    End With
    
    ' Add multiple lines of text to the shape
    With oShape.TextFrame2.TextRange
        .Text = "This is the first paragraph with the indent of 1 inch set to the first line. " & _
                "This indent is set by the paragraph style. No paragraph inline style is applied. " & _
                "These sentences are used to add lines for demonstrative purposes. " & _
                "These sentences are used to add lines for demonstrative purposes."
    End With
    
    ' Append a new paragraph displaying the class type
    With oShape.TextFrame2.TextRange
        .InsertAfter vbCrLf & "Class Type = " & sClassType
    End With
End Sub
```