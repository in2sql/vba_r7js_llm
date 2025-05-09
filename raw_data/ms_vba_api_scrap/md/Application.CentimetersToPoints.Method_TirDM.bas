Attribute VB_Name = "TirDM"
Sub dm2T()
    Application.ScreenUpdating = False

        With ThisWorkbook.Sheets("Tiras DM")
            .Range("C14, E16, C22, C24, C26, C28").Value = "x"
        End With

    Application.ScreenUpdating = True
End Sub
Sub dm1T()
Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("Tiras DM")
        .Range("C14, C16, G22, G24, G26, G28").Value = "x"
    End With
Application.EnableEvents = True
End Sub
Sub dmgT()
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("Tiras DM")
        .Range("C14, I16, G22, G24, G26, G28").Value = "x"
    End With
    Application.EnableEvents = True
End Sub

Sub dmReg()
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("Tiras DM")
        .Range("C18").Value = "x"
    End With
    Application.EnableEvents = True
End Sub
Sub dmNph()
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("Tiras DM")
        .Range("G18").Value = "x"
    End With
    Application.EnableEvents = True
End Sub
Sub dmGlic()
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets("Tiras DM")
        .Range("C20").Value = "x"
    End With
    Application.EnableEvents = True
End Sub
Sub ImprimirTiraDM()
Application.ScreenUpdating = False
    With ActiveSheet.PageSetup
        .PrintArea = "B3:M32"
        .PaperSize = xlPaperA4
        .Zoom = 105
        .LeftMargin = Application.CentimetersToPoints(0.6)
        .RightMargin = Application.CentimetersToPoints(0.6)
        .CenterHorizontally = True
        .CenterVertically = True
    End With
    ActiveSheet.PrintOut
Application.EnableEvents = True
End Sub

Sub NomeTira()
    Application.ScreenUpdating = False
     Sheets("Tiras DM").Range("C11").Value = Sheets("Receitas").Range("E14").Value
    Application.EnableEvents = True
End Sub
Sub limpatiras()
    Application.ScreenUpdating = False
        With ThisWorkbook.Sheets("Tiras DM")
            .Range("C11:J11, C12:J12, C13:C28, E13:E28, G13:G28, I13:I16").ClearContents
        End With
    Application.EnableEvents = True
End Sub








