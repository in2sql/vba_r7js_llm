# TextFrame2.MarginLeft Property (Excel)

## Business Description
The `MarginLeft` property of `TextFrame2` in Excel allows you to adjust the space between the left edge of a text frame and the left boundary of the shape that contains the text. This helps ensure your text appears visually balanced and is not too close to the shape's edge, improving readability and presentation.

## Behavior
- **Get or Set Value**: You can read or change the left margin (in points) of the text inside a shape.
- **Unit**: The value is measured in points.
- **Applies To**: Only applies to shapes that support the `TextFrame2` object (such as text boxes, WordArt, or shapes with text).
- **Version**: Available starting in Excel 2007.

## Example Usage
```vba
' Set the left margin of the text frame in a shape to 10 points
ActiveSheet.Shapes(1).TextFrame2.MarginLeft = 10

' Read the current left margin of the text frame
Dim margin As Single
margin = ActiveSheet.Shapes(1).TextFrame2.MarginLeft
```

**Tip:** Adjusting the margin can help prevent text from appearing cramped against the edge of shapes, making your reports and dashboards more professional.
