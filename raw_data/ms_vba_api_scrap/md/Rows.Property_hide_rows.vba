Sub Hide_rows()

' Hide specific row.

    Rows("8:28").Select
    Range("B28").Activate
    Selection.EntireRow.Hidden = False
    
    Rows("20:27").Select
    Selection.EntireRow.Hidden = True
    Range("C3").Select
End Sub
