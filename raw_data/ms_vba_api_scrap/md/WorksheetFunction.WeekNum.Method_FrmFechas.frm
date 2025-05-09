VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmFechas 
   Caption         =   "Formulario Ingresar Fecha"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "FrmFechas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FrmFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub btnConfirmar_Click()
''Generico esto funciona
'
''enero b5
''febrero j5
''marzo r5
''abril b14
''mayo j14
''junio r14
''julio b25
''agosto j25
''septiembre r25
''octubre b34
''noviembre j34
''diciembre r34
'
'Dim calenadrio As Range
'Dim fecha As Date
'Dim primerdiadelmes As Date
'Dim ultimodiadelmes As Date
'Dim dia As Integer
'Dim mes As Integer
'Dim diadelasemana As Integer
'Dim semanadelmes As Long
'Dim diassumar As Integer
'
'fecha = FrmFechas.txtFecha.Value
'primerdiadelmes = FrmFechas.txtFecha.Value
''Range("b2:x2").Value = fecha
''Range("b22:x22").Value = fecha
'calendario = Range("r14").Offset(0, 0).Select
'ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)
'
'
'While CDate(fecha) < CDate(ultimodiadelmes)
'fecha = DateAdd("d", diasasumar, fecha)
''MsgBox (fecha)
'diadelasemana = DatePart("w", fecha)
'semanadelmes = DatePart("ww", fecha) - DatePart("ww", primerdiadelmes) + 1
'' esto funciona
''WeekNumberFromDate = Int(((dt - StartDate) + 6) / 7) + Abs(Weekday(dt) = Weekday(StartDate))
''MsgBox (semanadelmes)
''MsgBox (semanadelmes)
'Range("r14").Offset(semanadelmes, diadelasemana - 1).Select
'Range("r14").Offset(semanadelmes, diadelasemana - 1).Value = fecha
'diasasumar = 1
'Wend
'
'End Sub

Private Sub btnConfirmar_Click()
'MES        CELDA
'enero      b5
'febrero    j5
'marzo      r5
'abril      b14
'mayo       j14
'junio      r14
'julio      b25
'agosto     j25
'septiembre r25
'octubre    b34
'noviembre  j34
'diciembre  r34

Dim calenadrio As Range
Dim celda As String
Dim fecha As Date
Dim primerdiadelmes As Date
Dim ultimodiadelmes As Date
Dim dia As Integer
Dim mes As Integer
Dim diadelasemana As Integer
Dim semanadelmes As Long
Dim diassumar As Integer
Dim buscar As New clsCalendario
celda = "b5"
fecha = FrmFechas.txtFecha.Value
primerdiadelmes = FrmFechas.txtFecha.Value
Range("b2:x2").Value = fecha
Range("b22:x22").Value = fecha
calendario = Range(celda).Offset(0, 0).Select
ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)
mes = DatePart("m", fecha)
'bucle meses
While mes < 13
'bucle dias
        While CDate(fecha) < CDate(ultimodiadelmes)
            fecha = DateAdd("d", diasasumar, fecha)
            diadelasemana = DatePart("w", fecha)
            semanadelmes = DatePart("ww", fecha) - DatePart("ww", primerdiadelmes) + 1
            Range(celda).Offset(semanadelmes, diadelasemana - 1).Select
            Range(celda).Offset(semanadelmes, diadelasemana - 1).Value = fecha
            diasasumar = 1
        Wend
        'reseteo variables
    fecha = DateAdd("m", 1, fecha)
    mes = DatePart("m", fecha)
    primerdiadelmes = DateSerial(Year(fecha), Month(fecha), 1)
    fecha = primerdiadelmes
    ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)
    diasasumar = 0
    Select Case mes
    Case 1
        celda = "b5"
    Case 2
        celda = "j5"
    Case 3
        celda = "r5"
    Case 4
        celda = "b14"
    Case 5
        celda = "j14"
    Case 6
        celda = "r14"
    Case 7
        celda = "b25"
    Case 8
        celda = "j25"
    Case 9
        celda = "r25"
    Case 10
        celda = "b34"
    Case 11
        celda = "j34"
    Case 12
        celda = "r34"
        While CDate(fecha) < CDate(ultimodiadelmes)
            fecha = DateAdd("d", diasasumar, fecha)
            diadelasemana = DatePart("w", fecha)
            semanadelmes = DatePart("ww", fecha) - DatePart("ww", primerdiadelmes) + 1
            Range(celda).Offset(semanadelmes, diadelasemana - 1).Select
            Range(celda).Offset(semanadelmes, diadelasemana - 1).Value = fecha
            diasasumar = 1
        Wend
        mes = 13
    End Select
    
Wend

    fecha = FrmFechas.txtFecha.Value
    Sheets("Diario").Select
    Range("A2").Select
    Range("A2").Value = fecha
    Range(Selection, Selection.End(xlDown)).Select
   
    Range("A2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("A2:A367")
    Range("A2:A367").Select
    Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"


MsgBox ("Finalizo Carga")

FrmFechas.Hide

End Sub



'Private Sub btnConfirmar_Click()
'-----ENERO----------
'Dim calenadrio As Range
'Dim fecha As Date
'Dim ultimodiadelmes As Date
'Dim dia As Integer
'Dim mes As Integer
'Dim diadelasemana As Integer
'Dim semanadelmes As Integer
'Dim diassumar As Integer
''esto esta bien para enero
'fecha = FrmFechas.txtFecha.Value
'Range("b2:x2").Value = fecha
'Range("b22:x22").Value = fecha
''dia = DatePart("d", fecha)
''mes = DatePart("m", fecha)
''diadelasemana = DatePart("w", fecha)
''semanadelmes = DatePart("ww", fecha)
'calendario = Range("b5").Offset(0, 0).Select
''MsgBox ("dia" & dia)
''MsgBox ("mes" & mes)
''MsgBox ("dia de la semana " & diadelasemana)
''MsgBox ("semana del mes " & semanadelmes)
''Range("b5").Offset(semanadelmes, diadelasemana - 1).Select
''Range("b5").Offset(semanadelmes, diadelasemana - 1).Value
''Range("b5").Offset(semanadelmes, diadelasemana - 1).Value = fecha
'
'
'ultimodiadelmes = Application.WorksheetFunction.EoMonth(fecha, 0)
'
'While fecha < ultimodiadelmes
'fecha = DateAdd("d", diasasumar, fecha)
''MsgBox (fecha)
'diadelasemana = DatePart("w", fecha)
'semanadelmes = DatePart("ww", fecha)
'Range("b5").Offset(semanadelmes, diadelasemana - 1).Select
'Range("b5").Offset(semanadelmes, diadelasemana - 1).Value = fecha
'diasasumar = 1
'Wend
'
'End Sub

'Private Sub btnConfirmar_Click()
'--FEBRERO--------
'' esto funciona
'Dim calenadrio As Range
'Dim fecha As Date
'Dim ultimodiadelmes As Date
'Dim dia As Integer
'Dim mes As Integer
'Dim diadelasemana As Integer
'Dim semanadelmes As Integer
'Dim diassumar As Integer
''esto esta bien para enero
'fecha = FrmFechas.txtFecha.Value
'Range("b2:x2").Value = fecha
'Range("b22:x22").Value = fecha
''dia = DatePart("d", fecha)
''mes = DatePart("m", fecha)
''diadelasemana = DatePart("w", fecha)
''semanadelmes = DatePart("ww", fecha)
'calendario = Range("j5").Offset(0, 0).Select
''MsgBox ("dia" & dia)
''MsgBox ("mes" & mes)
''MsgBox ("dia de la semana " & diadelasemana)
''MsgBox ("semana del mes " & semanadelmes)
''Range("b5").Offset(semanadelmes, diadelasemana - 1).Select
''Range("b5").Offset(semanadelmes, diadelasemana - 1).Value
''Range("b5").Offset(semanadelmes, diadelasemana - 1).Value = fecha
'
'
'ultimodiadelmes = Application.WorksheetFunction.EoMonth(fecha, 0)
'
'While fecha < ultimodiadelmes
'fecha = DateAdd("d", diasasumar, fecha)
''MsgBox (fecha)
'diadelasemana = DatePart("w", fecha)
'semanadelmes = Application.WorksheetFunction.WeekNum(fecha)
'MsgBox (semanadelmes)
'Range("j5").Offset(semanadelmes - 5, diadelasemana - 1).Select
'Range("j5").Offset(semanadelmes - 5, diadelasemana - 1).Value = fecha
'diasasumar = 1
'Wend
'
'End Sub

'Private Sub btnConfirmar_Click()
''---Marzo----
'Dim calenadrio As Range
'Dim fecha As Date
'Dim ultimodiadelmes As Date
'Dim dia As Integer
'Dim mes As Integer
'Dim diadelasemana As Integer
'Dim semanadelmes As Integer
'Dim diassumar As Integer
'
'fecha = FrmFechas.txtFecha.Value
''Range("b2:x2").Value = fecha
''Range("b22:x22").Value = fecha
'calendario = Range("r5").Offset(0, 0).Select
'ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)
'
'
'While CDate(fecha) < CDate(ultimodiadelmes)
'fecha = DateAdd("d", diasasumar, fecha)
''MsgBox (fecha)
'diadelasemana = DatePart("w", fecha)
'semanadelmes = DatePart("ww", fecha)
'' esto funciona
''WeekNumberFromDate = Int(((dt - StartDate) + 6) / 7) + Abs(Weekday(dt) = Weekday(StartDate))
''MsgBox (semanadelmes)
''MsgBox (semanadelmes)
'Range("r5").Offset(semanadelmes - 9, diadelasemana - 1).Select
'Range("r5").Offset(semanadelmes - 9, diadelasemana - 1).Value = fecha
'diasasumar = 1
'Wend
'
'End Sub

'Private Sub btnConfirmar_Click()
''---abril----
'Dim calenadrio As Range
'Dim fecha As Date
'Dim ultimodiadelmes As Date
'Dim dia As Integer
'Dim mes As Integer
'Dim diadelasemana As Integer
'Dim semanadelmes As Integer
'Dim diassumar As Integer
'
'fecha = FrmFechas.txtFecha.Value
''Range("b2:x2").Value = fecha
''Range("b22:x22").Value = fecha
'calendario = Range("b14").Offset(0, 0).Select
'ultimodiadelmes = Application.WorksheetFunction.EoMonth(CDate(fecha), 0)
'
'
'While CDate(fecha) < CDate(ultimodiadelmes)
'fecha = DateAdd("d", diasasumar, fecha)
''MsgBox (fecha)
'diadelasemana = DatePart("w", fecha)
'semanadelmes = DatePart("ww", fecha)
'' esto funciona
''WeekNumberFromDate = Int(((dt - StartDate) + 6) / 7) + Abs(Weekday(dt) = Weekday(StartDate))
''MsgBox (semanadelmes)
''MsgBox (semanadelmes)
'Range("b14").Offset(semanadelmes - 13, diadelasemana - 1).Select
'Range("b14").Offset(semanadelmes - 13, diadelasemana - 1).Value = fecha
'diasasumar = 1
'Wend
'
'End Sub


