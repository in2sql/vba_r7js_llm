# LineFormat InsetPen Property

## Business Description
Returns or sets whether lines are drawn inside the specified shape's boundaries. Read/write

## Behavior
Returns or sets whether lines are drawn inside the specified shape's boundaries. Read/write

## Example Usage
```vba
Dim shpNew As Shape 
 
With ActiveSheet.Shapes 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=150, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen= msoTrue 
 End With 
 
 Set shpNew = .AddShape(Type:=msoShapeRectangle, _ 
 Left:=200, Top:=300, Width:=150, Height:=100) 
 With shpNew.Line 
 .Weight = 24 
 .InsetPen= msoFalse 
 End With 
End With
```