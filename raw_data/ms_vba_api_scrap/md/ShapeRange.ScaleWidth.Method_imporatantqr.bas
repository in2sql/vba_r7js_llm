Attribute VB_Name = "Module2"
Sub addQRCodes()

    For Each cell In Selection

    cell.Offset(0, 1).Select
    filepath = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" & WorksheetFunction.EncodeURL(cell.Value)
    With ActiveSheet.Pictures.Insert(filepath)
        .ShapeRange.ScaleWidth 0.85, msoFalse, msoScaleFromTopLeft
        .ShapeRange.ScaleHeight 0.85, msoFalse, msoScaleFromTopLeft
    End With
    
    Next cell
    
End Sub

Sub removeQRCodes()
    For Each pic In ActiveSheet.Pictures
        pic.Delete
    Next pic
End Sub

