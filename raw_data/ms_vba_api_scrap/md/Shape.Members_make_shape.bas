Attribute VB_Name = "Module1"
Sub make_shape()
    Dim new_shape As Shape
    Dim point_x As Integer
    Dim point_y As Integer
    Dim width As Integer
    Dim height As Integer
    
    x = Selection.Left
    y = Selection.Top
    
    width = 100
    height = 70
    
    Set new_shape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, x, y, width, height)
    
    new_shape.TextFrame2.TextRange.Text = "test"
    With new_shape.TextFrame2.TextRange.Font
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Size = 20
    End With
    new_shape.TextFrame2.HorizontalAnchor = msoAnchorCenter
    new_shape.OnAction = "'test ""5""'"
    new_shape.Name = "01"
    
    new_shape.Fill.ForeColor.RGB = RGB(200, 150, 150)
    new_shape.Line.ForeColor.RGB = RGB(255, 0, 0)
    new_shape.Line.Weight = 4
    
    Set new_shape = Nothing
    
End Sub

Sub test(x As Integer)
    MsgBox x

End Sub
