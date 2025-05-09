# Description / Описание

**English:**  
This code sets the left indentation of a paragraph within a shape on the active worksheet. It creates a shape with specific dimensions and fill properties, adds text to the paragraph, and applies an indentation of 2 inches.

**Русский:**  
Этот код устанавливает левый отступ абзаца внутри фигуры на активном рабочем листе. Он создает фигуру с определенными размерами и свойствами заливки, добавляет текст в абзац и применяет отступ в 2 дюйма.

## VBA Code

```vba
' This VBA code sets the left indentation of a paragraph within a shape on the active worksheet

Sub SetParagraphIndentation()
    Dim oWorksheet As Object
    Dim oFill As Object
    Dim oStroke As Object
    Dim oShape As Object
    Dim oDocContent As Object
    Dim oParagraph As Object
    Dim oParaPr As Object
    
    ' Get the active sheet
    Set oWorksheet = Api.GetActiveSheet()
    
    ' Create a solid fill with RGB color (255, 111, 61)
    Set oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61))
    
    ' Create a stroke with no fill
    Set oStroke = Api.CreateStroke(0, Api.CreateNoFill())
    
    ' Add a shape to the worksheet with specified parameters
    Set oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000)
    
    ' Get the content of the shape
    Set oDocContent = oShape.GetContent()
    
    ' Get the first paragraph in the content
    Set oParagraph = oDocContent.GetElement(0)
    
    ' Get the paragraph properties
    Set oParaPr = oParagraph.GetParaPr()
    
    ' Set the left indentation to 2880 (2 inches)
    oParaPr.SetIndLeft(2880)
    
    ' Add text to the paragraph
    oParagraph.AddText "This is the first paragraph with the indent of 2 inches set to it. "
    oParagraph.AddText "This indent is set by the paragraph style. No paragraph inline style is applied. "
    oParagraph.AddText "These sentences are used to add lines for demonstrative purposes. "
End Sub
```

## OnlyOffice JS Code

```javascript
// This OnlyOffice JS code sets the left indentation of a paragraph within a shape on the active worksheet

// Get the active worksheet
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph in the content
var oParagraph = oDocContent.GetElement(0);

// Get the paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set the left indentation to 2880 (2 inches)
oParaPr.SetIndLeft(2880);

// Add text to the paragraph
oParagraph.AddText("This is the first paragraph with the indent of 2 inches set to it. ");
oParagraph.AddText("This indent is set by the paragraph style. No paragraph inline style is applied. ");
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. ");
```