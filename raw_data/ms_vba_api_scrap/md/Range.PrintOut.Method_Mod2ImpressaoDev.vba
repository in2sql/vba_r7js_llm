Sub ImprimirDEV()
    Dim ws As Worksheet
    Dim printRange As Range

    ' Definir a planilha e o intervalo a serem impressos
    Set ws = ThisWorkbook.Sheets("COMPROVANTE DEVOLUÇÃO")
    Set printRange = ws.Range("B1:J46")

    ' Exibir a prévia antes de imprimir
    If Not Previa() Then Exit Sub

    ' Definir a configuração da impressora
    Application.Dialogs(xlDialogPrinterSetup).Show

    ' Imprimir o intervalo selecionado
    printRange.PrintOut Copies:=1, Collate:=True
End Sub

Function Previa() As Boolean
    Dim Resp As VbMsgBoxResult

    Resp = MsgBox("Deseja Imprimir o comprovante do técnico selecionado?", vbYesNo)

    If Resp = vbNo Then
        Previa = False
    Else
        Previa = True
    End If
End Function

