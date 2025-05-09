Set myDocument = Worksheets(1) 
Set rac = myDocument.Shapes.AddShape(msoShapeRightArrowCallout, _ 
 10, 10, 250, 190) 
With rac.Adjustments 
 .Item(1) = 0.5 'adjusts width of text box 
 .Item(2) = 0.15 'adjusts width of arrow head 
 .Item(3) = 0.8 'adjusts length of arrow head 
 .Item(4) = 0.4 'adjusts width of arrow neck 
End With