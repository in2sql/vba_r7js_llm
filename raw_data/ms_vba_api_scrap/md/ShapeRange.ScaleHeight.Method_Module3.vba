'//////////////////////////////////////////////////////////////////////////
'// Module3: Image
'//////////////////////////////////////////////////////////////////////////

'
' ImageSixtyPercent Macro
' Keyboard Shortcut: Ctrl+Shift+Z
Sub ImageSixtyPercent()

    Select Case TypeName(Selection)
    '画像が選択済み
    Case "Picture"
        Selection.ShapeRange.ScaleHeight 0.6, msoFalse, msoScaleFromTopLeft
        Application.CommandBars("Format Object").Visible = False
    Case Else
    
    End Select

End Sub

