Attribute VB_Name = "ModuloAux"
'paso lAs blcemi.guardias separadas para poder reutilizar la funcion
'Public Sub LlenarHoja(pHoja, pEmpleado As blcemi.Empleado, pGuardias As blcemi.guardiaManager, pHojaResumen, fila As Integer)
'    Dim g As blcemi.guardia
'    pHoja.Name = pEmpleado.NombreCompleto
'    pHoja.Cells(1, 1) = "Dia"
'    pHoja.Cells(1, 2) = "Valor de Guardia"
'    pHoja.Cells(1, 3) = "Coseguro"
'    pHoja.Cells(1, 4) = "Plus"
'    pHoja.Cells(1, 5) = "Subtotal"
'
'    Dim i As Integer
'    i = 2
'
'    For Each g In pGuardias
'        pHoja.Cells(i, 1) = g.fecha
'        pHoja.Cells(i, 2) = g.Monto
'        pHoja.Cells(i, 3) = g.Adelanto
'        pHoja.Cells(i, 4) = g.Plus
'        pHoja.Cells(i, 5) = g.GetSaldo
'        i = i + 1
'    Next
'
'    pHoja.Cells(i, 2).formula = "=sum(B1:B" + Trim(Str(i - 1)) + ")"
'    pHoja.Cells(i, 3).formula = "=sum(C1:C" + Trim(Str(i - 1)) + ")"
'    pHoja.Cells(i, 4).formula = "=sum(D1:D" + Trim(Str(i - 1)) + ")"
'    pHoja.Cells(i, 5).formula = "=sum(E1:E" + Trim(Str(i - 1)) + ")"
'
'    'colocamos en negrita los enbezados en la hoja
'    pHoja.Rows(1).Font.Bold = True
'    pHoja.Rows(i).Font.Bold = True
'
'    'Autoajustamos
'    pHoja.Columns("A:Z").AutoFit
'
'    'agregamos info al resumen
'    If Not pHojaResumen Is Nothing Then
'        pHojaResumen.Cells(fila, 2) = pEmpleado.NombreCompleto
'        pHojaResumen.Cells(fila, 3).formula = "='" + pEmpleado.NombreCompleto + "'!B" + Trim(Str(i))
'        pHojaResumen.Cells(fila, 4).formula = "='" + pEmpleado.NombreCompleto + "'!C" + Trim(Str(i))
'        pHojaResumen.Cells(fila, 5).formula = "='" + pEmpleado.NombreCompleto + "'!D" + Trim(Str(i))
'        pHojaResumen.Cells(fila, 6).formula = "='" + pEmpleado.NombreCompleto + "'!E" + Trim(Str(i))
'    End If
'
'End Sub

Public Enum eXLBorderWidth
    eXLThick = -4138
    eXLThin = 2
End Enum

'para usarlo en ayuda puntual
Public Sub VerManual(pagina As Integer)
Dim wApp As Object
Dim doc As Object
Set wApp = CreateObject("word.application")
If Not wApp Is Nothing Then
    wApp.Visible = True
    wApp.Documents.open Chr(32) + APh + "..\manual\manual de usuario.doc" + Chr(32), , True
    wApp.Selection.GoTo 1, 1, pagina
End If
End Sub


Public Sub UnirYCentrar(rango, Optional pBorde As Boolean = True)
    rango.merge
    rango.HorizontalAlignment = -4108
    If pBorde Then Bordes rango, eXLThick
End Sub


Public Sub Bordes(range, Optional borde As eXLBorderWidth = eXLBorderWidth.eXLThin, Optional incluirInteriores As Boolean = False)
'range.Select
range.Borders(5).LineStyle = -4142 'xlNonexlDiagonalDown
range.Borders(6).LineStyle = -4142 ' xlNonexlDiagonalUp
With range.Borders(7) 'xlEdgeLeft)
.LineStyle = 1 'xlContinuous
.Weight = borde
.ColorIndex = -4105 'xlAutomatic
End With
With range.Borders(8) 'xlEdgeTop)
.LineStyle = 1 'xlContinuous
.Weight = borde
.ColorIndex = -4105 'xlAutomatic
End With
With range.Borders(9) 'xlEdgeBottom)
.LineStyle = 1 'xlContinuous
.Weight = borde
.ColorIndex = -4105 'xlAutomatic
End With
With range.Borders(10) 'xlEdgeRight)
.LineStyle = 1 'xlContinuous
.Weight = borde
.ColorIndex = -4105 'xlAutomatic
End With
If incluirInteriores Then
'    range.Borders(11).LineStyle = 1 ' xlNone 'xlInsideVertical=11
'    range.Borders(12).LineStyle = 1
    range.Borders(11).Weight = eXLThin
    range.Borders(12).Weight = eXLThin
Else
    range.Borders(11).LineStyle = -4142 ' xlNone 'xlInsideVertical=11
    range.Borders(12).LineStyle = -4142 'xlNone 'xlInsideHorizontal=12
End If
End Sub


Public Function GetParametros(pPath As String) As LParameterManager
    On Error GoTo errman
    Dim params As New LParameterManager
    Dim Nombre As String
    Dim Tipo As String
    Dim Descripcion As String
    Dim I As Integer
    I = 0
    Do
        Nombre = Leer_Ini(pPath, "Parametros", "Nombre" + Trim(Str(I)), "")
        Tipo = Leer_Ini(pPath, "Parametros", "Tipo" + Trim(Str(I)), "")
        Descripcion = Leer_Ini(pPath, "Parametros", "Descripcion" + Trim(Str(I)), "")

        If Nombre <> "" Then
            params.Add Nombre, Tipo, Descripcion
        Else
            Exit Do
        End If
        I = I + 1
    Loop
    Set GetParametros = params
    Exit Function
errman:
    Set GetParametros = Nothing
End Function


