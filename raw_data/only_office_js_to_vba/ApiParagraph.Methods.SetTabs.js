```javascript
// This code sets custom tab stops in a paragraph.
// Этот код устанавливает пользовательские табуляции в абзаце.

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
    150 * 36000,
    70 * 36000,
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

// Set custom tab stops: 1440, 2880, 4320 units with alignment left, center, right
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

```vba
' This code sets custom tab stops in a paragraph.
' Этот код устанавливает пользовательские табуляции в абзаце.

' VBA Code Equivalent

Sub SetCustomTabStops()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' Get the active worksheet
    
    ' Add a shape to the worksheet
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape( _
        Type:=msoShapeFlowchartDatabase, _
        Left:=150, Top:=70, Width:=2, Height:=3)
    
    ' Set fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
    
    ' Remove the stroke
    With shp.Line
        .Visible = msoFalse
    End With
    
    ' Access the text frame of the shape
    With shp.TextFrame
        .HorizontalAlignment = xlHAlignLeft
        .Characters.Text = "Custom tab - 1 inch left" & vbCrLf & _
                          "Custom tab - 2 inches center" & vbCrLf & _
                          "Custom tab - 3 inches right"
        
        ' Set custom tab stops
        .Characters(1, 25).ParagraphFormat.TabStops.ClearAll
        .Characters(1, 25).ParagraphFormat.TabStops.Add Position:=72 ' 1 inch
        .Characters(1, 25).ParagraphFormat.TabStops.Add Position:=144 ' 2 inches, center
        .Characters(1, 25).ParagraphFormat.TabStops.Add Position:=216 ' 3 inches, right
        
        ' Note: VBA does not support setting alignment for individual tab stops directly
        ' Advanced alignment would require additional formatting
    End With
End Sub
```