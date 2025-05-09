# Adjustments object (Excel)

## Description
This page is from the Excel VBA API reference. The content might be limited due to browser compatibility issues.

## Remarks
Each adjustment value represents one way that an adjustment handle can be adjusted. Because some adjustment handles can be adjusted in two waysâfor example, some handles can be adjusted both horizontally and verticallyâa shape can have more adjustment values than it has adjustment handles. A shape can have up to eight adjustments.

## Example
```vba
Set myDocument = Worksheets(1) 
Set rac = myDocument.Shapes.AddShape(msoShapeRightArrowCallout, _ 
 10, 10, 250, 190) 
With rac.Adjustments 
 .Item(1) = 0.5 'adjusts width of text box 
 .Item(2) = 0.15 'adjusts width of arrow head 
 .Item(3) = 0.8 'adjusts length of arrow head 
 .Item(4) = 0.4 'adjusts width of arrow neck 
End With
```

