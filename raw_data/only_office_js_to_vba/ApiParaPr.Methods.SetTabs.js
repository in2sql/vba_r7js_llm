## Description / Описание

**English:**  
This code adds a shape to the active worksheet with a specific fill color and no stroke. It sets custom tab stops for the paragraph within the shape and adds text with line breaks aligned according to the tab stops.

**Russian:**  
Этот код добавляет фигуру на активный лист с определенным цветом заливки и без обводки. Он устанавливает пользовательские табуляции для абзаца внутри фигуры и добавляет текст с разрывами строк, выровненными согласно табуляциям.

## VBA Code

```vba
' Adds a shape to the active worksheet with specified fill and no stroke,
' sets custom tab stops for the paragraph, and adds aligned text with line breaks.

Sub AddCustomShape()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Define RGB color
    Dim fillColor As Long
    fillColor = RGB(255, 111, 61)
    
    ' Add shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 150, 70, 150, 70) ' Adjust sizes as needed
    
    ' Set fill color
    shp.Fill.ForeColor.RGB = fillColor
    
    ' No line (stroke)
    shp.Line.Visible = msoFalse
    
    ' Add text to the shape
    With shp.TextFrame2
        ' Set paragraph alignment
        Dim para As ParagraphFormat2
        Set para = .TextRange.ParagraphFormat
        ' Set tab stops at 1 inch, 2 inches, 3 inches with different alignments
        para.TabStops.ClearAll
        para.TabStops.Add Type:=msoTabStopLeft, Position:=1440, Alignment:=msoTabStopLeft
        para.TabStops.Add Type:=msoTabStopCenter, Position:=2880
        para.TabStops.Add Type:=msoTabStopRight, Position:=4320
        
        ' Add text with tabs and line breaks
        .TextRange.Text = vbTab & "Custom tab - 1 inch left" & vbCrLf & _
                          vbTab & vbTab & "Custom tab - 2 inches center" & vbCrLf & _
                          vbTab & vbTab & vbTab & "Custom tab - 3 inches right"
    End With
End Sub
```

## OnlyOffice JS Code

```javascript
// This example sets a sequence of custom tab stops which will be used for any tab characters in the paragraph.
var oWorksheet = Api.GetActiveSheet();

// Create a solid fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));

// Create a stroke with width 0 and no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());

// Add a shape to the worksheet with specified parameters
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 150 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);

// Get the content of the shape
var oDocContent = oShape.GetContent();

// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);

// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();

// Set custom tab stops at positions 1440, 2880, 4320 with alignments left, center, right
oParaPr.SetTabs([1440, 2880, 4320], ["left", "center", "right"]);

// Add tab stops and text with line breaks
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 1 inch left");
oParagraph.AddLineBreak();

oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 2 inches center");
oParagraph.AddLineBreak();

oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 3 inches right");
```