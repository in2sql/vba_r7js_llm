```vba
' This code sets a sequence of custom tab stops used for any tab characters in the paragraph.
' Этот код устанавливает последовательность пользовательских табуляторов, используемых для любых символов табуляции в параграфе.

Sub SetCustomTabStops()
    ' Get the active worksheet
    Dim oWorksheet As Worksheet
    Set oWorksheet = ThisWorkbook.ActiveSheet
    
    ' Add a shape to the worksheet
    Dim oShape As Shape
    Set oShape = oWorksheet.Shapes.AddShape(msoShapeFlowchartProcess, 150, 70, 150, 70)
    
    ' Set the fill color to RGB(255, 111, 61)
    With oShape.Fill
        .Solid
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
    
    ' Set the line style to have no fill
    With oShape.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    With oShape.TextFrame2
        ' Clear existing paragraphs
        .TextRange.Text = ""
        
        ' Add a new paragraph
        Dim oParagraph As TextRange2
        Set oParagraph = .TextRange.Paragraphs.Add
        
        ' Set custom tab stops at 1 inch, 2 inches, and 3 inches
        With oParagraph.ParagraphFormat.TabStops
            .ClearAll
            .Add Position:=72, Alignment:=msoTabStopLeft ' 1 inch
            .Add Position:=144, Alignment:=msoTabStopCenter ' 2 inches
            .Add Position:=216, Alignment:=msoTabStopRight ' 3 inches
        End With
        
        ' Add text with tabs and line breaks
        oParagraph.Text = vbTab & "Custom tab - 1 inch left" & vbCrLf & _
                          vbTab & vbTab & "Custom tab - 2 inches center" & vbCrLf & _
                          vbTab & vbTab & vbTab & "Custom tab - 3 inches right"
    End With
End Sub
```

```javascript
// This code sets a sequence of custom tab stops used for any tab characters in the paragraph.
// Этот код устанавливает последовательность пользовательских табуляторов, используемых для любых символов табуляции в параграфе.

var oWorksheet = Api.GetActiveSheet();
// Create a fill with RGB color (255, 111, 61)
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
// Create a stroke with no fill
var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
// Add a flowchart process shape to the worksheet
var oShape = oWorksheet.AddShape("flowChartProcess", 150 * 36000, 70 * 36000, oFill, oStroke, 150 * 36000, 70 * 36000, 150 * 36000, 70 * 36000);
// Get the document content of the shape
var oDocContent = oShape.GetContent();
// Get the first paragraph in the document content
var oParagraph = oDocContent.GetElement(0);
// Set custom tab stops at positions 1440, 2880, 4320 with alignments left, center, right
oParagraph.SetTabs([1440, 2880, 4320], ["left", "center", "right"]);
// Add a tab stop and text for the first tab
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 1 inch left");
// Add a line break
oParagraph.AddLineBreak();
// Add two tab stops and text for the second tab
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 2 inches center");
// Add a line break
oParagraph.AddLineBreak();
// Add three tab stops and text for the third tab
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddTabStop();
oParagraph.AddText("Custom tab - 3 inches right");
```