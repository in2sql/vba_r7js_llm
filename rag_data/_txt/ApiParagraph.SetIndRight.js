```plaintext
// This code sets the paragraph right side indentation and alignment.
// Этот код устанавливает отступ справа и выравнивание абзаца.
```

```vba
' VBA Code to set paragraph right indentation and alignment in a shape

Sub SetParagraphIndentationAndAlignment()
    ' Declare variables
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtFrame As TextFrame2
    Dim txtRange As TextRange2
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Add a shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStoredData, 120, 70, 300, 150)
    
    ' Set the fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
    
    ' Remove the line (stroke)
    shp.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    Set txtFrame = shp.TextFrame2
    
    ' Add first paragraph with right alignment and right indentation
    With txtFrame.TextRange
        .Text = "This is a paragraph with the right offset of 2 inches set to it. " & _
                "We also aligned the text in it by the right side. " & _
                "This sentence is used to add lines for demonstrative purposes."
        .ParagraphFormat.Alignment = msoAlignRight
        .ParagraphFormat.RightIndent = 2880 ' Points
    End With
    
    ' Add a second paragraph without any indentation
    With txtFrame.TextRange
        .InsertAfter vbCrLf & "This is a paragraph without any offset set to it. " & _
                     "These sentences are used to add lines for demonstrative purposes."
        .Characters(Start:=Len(.Text) - Len("These sentences are used to add lines for demonstrative purposes.") + 1, Length:=Len("These sentences are used to add lines for demonstrative purposes.")).ParagraphFormat.RightIndent = 0
        .Characters(Start:=Len(.Text) - Len("These sentences are used to add lines for demonstrative purposes.") + 1, Length:=Len("These sentences are used to add lines for demonstrative purposes.")).ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub
```

```javascript
// This code sets the paragraph right side indentation and alignment.
// Этот код устанавливает отступ справа и выравнивание абзаца.

// VBA Code to set paragraph right indentation and alignment in a shape

Sub SetParagraphIndentationAndAlignment()
    ' Declare variables
    Dim ws As Worksheet
    Dim shp As Shape
    Dim txtFrame As TextFrame2
    Dim txtRange As TextRange2
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Add a shape to the worksheet
    Set shp = ws.Shapes.AddShape(msoShapeFlowchartStoredData, 120, 70, 300, 150)
    
    ' Set the fill color to RGB(255, 111, 61)
    With shp.Fill
        .Visible = msoTrue
        .Solid
        .ForeColor.RGB = RGB(255, 111, 61)
    End With
    
    ' Remove the line (stroke)
    shp.Line.Visible = msoFalse
    
    ' Access the text frame of the shape
    Set txtFrame = shp.TextFrame2
    
    ' Add first paragraph with right alignment and right indentation
    With txtFrame.TextRange
        .Text = "This is a paragraph with the right offset of 2 inches set to it. " & _
                "We also aligned the text in it by the right side. " & _
                "This sentence is used to add lines for demonstrative purposes."
        .ParagraphFormat.Alignment = msoAlignRight
        .ParagraphFormat.RightIndent = 2880 ' Points
    End With
    
    ' Add a second paragraph without any indentation
    With txtFrame.TextRange
        .InsertAfter vbCrLf & "This is a paragraph without any offset set to it. " & _
                     "These sentences are used to add lines for demonstrative purposes."
        .Characters(Start:=Len(.Text) - Len("These sentences are used to add lines for demonstrative purposes.") + 1, Length:=Len("These sentences are used to add lines for demonstrative purposes.")).ParagraphFormat.RightIndent = 0
        .Characters(Start:=Len(.Text) - Len("These sentences are used to add lines for demonstrative purposes.") + 1, Length:=Len("These sentences are used to add lines for demonstrative purposes.")).ParagraphFormat.Alignment = msoAlignLeft
    End With
End Sub
```

```javascript
// This code sets the paragraph right side indentation and alignment.
// Этот код устанавливает отступ справа и выравнивание абзаца.

// OnlyOffice JS Code

var oWorksheet = Api.GetActiveSheet(); // Get the active sheet
var oFill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61)); // Create a solid fill with RGB color
var oStroke = Api.CreateStroke(0, Api.CreateNoFill()); // Create a stroke with no fill
var oShape = oWorksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, oFill, oStroke, 0, 2 * 36000, 0, 3 * 36000); // Add a shape to the worksheet
var oDocContent = oShape.GetContent(); // Get the content of the shape
var oParagraph = oDocContent.GetElement(0); // Get the first paragraph
oParagraph.AddText("This is a paragraph with the right offset of 2 inches set to it. "); // Add text to the paragraph
oParagraph.AddText("We also aligned the text in it by the right side. "); // Add more text
oParagraph.AddText("This sentence is used to add lines for demonstrative purposes."); // Add more text
oParagraph.SetJc("right"); // Set justification to right
oParagraph.SetIndRight(2880); // Set right indentation
oParagraph = Api.CreateParagraph(); // Create a new paragraph
oParagraph.AddText("This is a paragraph without any offset set to it. "); // Add text to new paragraph
oParagraph.AddText("These sentences are used to add lines for demonstrative purposes. "); // Add more text
oDocContent.Push(oParagraph); // Add the new paragraph to the document content
```