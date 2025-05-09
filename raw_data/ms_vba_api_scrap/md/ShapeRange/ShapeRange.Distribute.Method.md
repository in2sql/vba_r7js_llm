# ShapeRange Distribute Method

## Business Description
Horizontally or vertically distributes the shapes in the specified range of shapes.

## Behavior
Horizontally or vertically distributes the shapes in the specified range of shapes.

## Example Usage
```vba
Set myDocument = Worksheets(1) 
With myDocument.Shapes 
    numShapes = .Count 
    If numShapes > 1 Then 
        numAutoShapes = 0 
        ReDim autoShpArray(1 To numShapes) 
        For i = 1 To numShapes 
            If .Item(i).Type = msoAutoShape Then 
                numAutoShapes = numAutoShapes + 1 
                autoShpArray(numAutoShapes) = .Item(i).Name 
            End If 
        Next 
        If numAutoShapes > 1 Then 
            ReDim Preserve autoShpArray(1 To numAutoShapes) 
            Set asRange = .Range(autoShpArray) 
            asRange.DistributemsoDistributeHorizontally, False 
        End If 
    End If 
End With
```