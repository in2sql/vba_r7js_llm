# ShapeNodes Insert Method

## Business Description
Inserts a node into a freeform shape.

## Behavior
Inserts a node into a freeform shape.

## Example Usage
```vba
Sub InsertShapeNode() 
    ActiveSheet.Shapes(3).Select 
    With Selection.ShapeRange 
        If .Type = msoFreeform Then 
            .Nodes.Insert_ 
                Index:=3, SegmentType:=msoSegmentCurve, _ 
                EditingType:=msoEditingSymmetric, X1:=35, Y1:=100 
            .Fill.ForeColor.RGB = RGB(0, 0, 200) 
            .Fill.Visible = msoTrue 
        Else 
            MsgBox "This shape is not a Freeform object." 
        End If 
    End With 
End Sub
```