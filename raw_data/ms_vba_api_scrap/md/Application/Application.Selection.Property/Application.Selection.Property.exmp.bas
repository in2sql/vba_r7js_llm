Sub TestSelection(  )
    Dim str As String
    Select Case TypeName(Selection)
    Case "Nothing"
        str = "No selection made."
    Case "Range"
        str = "You selected the range: " & Selection.Address
    Case "Picture"
        str = "You selected a picture."
    Case Else
        str = "You selected a " & TypeName(Selection) & "."
    End Select
    MsgBox str
End Sub