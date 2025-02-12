**Description / Описание**

This code sets a sequence of custom tab stops for any tab characters in a paragraph within an Excel worksheet.  
Этот код устанавливает последовательность пользовательских табуляций для любых символов табуляции в абзаце в рабочем листе Excel.

```vba
' VBA Code to set custom tab stops in Excel
Sub SetCustomTabStops()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Create a shape with specified dimensions and fill
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartOnlineStorage, 150, 70, 150, 70)
    
    ' Access the text frame of the shape
    With shp.TextFrame2.TextRange.ParagraphFormat
        ' Set tab stops at 1 inch, 2 inches, and 3 inches with left, center, right alignment
        .TabStops.ClearAll
        .TabStops.Add Position:=72, Alignment:=msoTabStopLeft
        .TabStops.Add Position:=144, Alignment:=msoTabStopCenter
        .TabStops.Add Position:=216, Alignment:=msoTabStopRight
    End With
    
    ' Add text with tabs
    With shp.TextFrame2.TextRange
        .Text = vbTab & "Custom tab - 1 inch left" & vbCrLf & vbTab & vbTab & "Custom tab - 2 inches center" & vbCrLf & vbTab & vbTab & vbTab & "Custom tab - 3 inches right"
    End With
End Sub
```

```javascript
// JavaScript Code to set custom tab stops in OnlyOffice
// This example sets a sequence of custom tab stops which will be used for any tab characters in the paragraph.
var oWorksheet = Api.GetActiveSheet();
// Create a solid fill with RGB color
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add a shape to the worksheet with specified dimensions and fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 150 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000);
// Get the content of the shape
var oDocContent = oShape.GetContent();
// Get the first paragraph element
var oParagraph = oDocContent.GetElement(0);
// Get paragraph properties
var oParaPr = oParagraph.GetParaPr();
// Set tab stops at positions 1440, 2880, 4320 with left, center, right alignment
oParaPr.SetTabs([1440, 2880, 4320], ["left", "center", "right"]);
// Add tab and text for first tab stop
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 1 inch left");
oParagraph.AddLineBreak();
// Add two tab stops and text for second tab stop
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 2 inches center");
oParagraph.AddLineBreak();
// Add three tab stops and text for third tab stop
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 3 inches right");
```