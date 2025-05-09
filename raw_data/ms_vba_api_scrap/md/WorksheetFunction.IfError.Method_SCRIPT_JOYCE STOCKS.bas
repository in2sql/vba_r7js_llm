Attribute VB_Name = "Module1"
Sub stocks()
Dim cantidadhojas As Integer
Dim valoropen As Double
Dim valorcierre As Double
Dim sumadestocks As Double
Dim contadorayuda As Double
sumadestocks = 0
Dim GREATINCRE As Double
Dim GREATDECRE As Double
Dim GREATTOTAL As Double
Dim TICKERINCRE As String
Dim TICKERDECRE As String
Dim TICKERGREAT As String
GREATINCRE = 0
GREATDECRE = 0
GREATTOTAL = 0


cantidadhojas = Sheets.Count()

    For hoja = 1 To cantidadhojas
       Worksheets(hoja).Activate
       With ActiveSheet
       UltCelda = .Cells(.Rows.Count, "a").End(xlUp).Row 'busca la ultima celda usada'
       Range("I1") = "Ticker"
       Range("j1") = "Yearly Change"
       Range("k1") = "Percent Change"
       Range("l1") = "Total Stock Value"
       Range("p1") = "Ticker"
       Range("q1") = "Value"
       Range("o2") = "GREATEST % INCREASE"
       Range("o3") = "GREATEST % DECREASE"
       Range("o4") = "GREATEST TOTAL VOLUME"
       End With
       Range(Cells(2, 1), Cells(UltCelda, 1)).Copy   'va a copiar la columna para encontrar los tickers unicos'
       Range("I2").PasteSpecial Paste:=xlPasteValues   'pega valores en columna para resumen'
       Range(Cells(2, 9), Cells(UltCelda, 9)).RemoveDuplicates Columns:=1, Header:=xlNo 'remuevo duplicados'
       With ActiveSheet
       Cantidadtickers = .Cells(.Rows.Count, "i").End(xlUp).Row 'cantidad de tickers unicos o ultima celda'
       End With
           
           
        For TICKER = 2 To Cantidadtickers 'loop para la busqueda de cada tiker'
           contadorayuda = 1
                For i = 2 To UltCelda 'loop para la buscada de cada registro'
                 If Cells(i, 1) = Cells(TICKER, 9) Then
                 sumadestocks = sumadestocks + Cells(i, 7)
                    If contadorayuda = 1 Then
                    valoropen = Cells(i, 3)
                    End If
                contadorayuda = contadorayuda + 1  'AUXILIAR PARA SETEAR LA ULTIMA LINEA QUE CONTIENE EL VALOR DE TICKER CON LA LETRA'
                 Else:
                 valorcierre = Cells(i - 1, 6)
                 Cells(TICKER, 12) = sumadestocks
                 Cells(TICKER, 10) = WorksheetFunction.IfError(valorcierre - valoropen, 0)
                 
                 
                   If Cells(TICKER, 10) < 0 Then 'color de la cenlda'
                   Cells(TICKER, 10).Interior.ColorIndex = 3
                   Else
                   Cells(TICKER, 10).Interior.ColorIndex = 4
                   End If
                   
                   
                   If valoropen = 0 Then 'evita errores cuando el valor es 0 al inicio)
                   Cells(TICKER, 11) = 0
                   Else:
                   Cells(TICKER, 11) = (valorcierre / valoropen) - 1
                   Cells(TICKER, 11).NumberFormat = "0.00%"
                   End If
                   
                   If Cells(TICKER, 12) > GREATTOTAL Then
                   TICKERGREAT = Cells(TICKER, 9)
                   GREATTOTAL = Cells(TICKER, 12)
                   End If
                   
                   If Cells(TICKER, 11) > GREATINCRE Then
                   TICKERINCRE = Cells(TICKER, 9)
                   GREATINCRE = Cells(TICKER, 11)
                   End If
                   
                   If Cells(TICKER, 11) < GREATDECRE Then
                   TICKERDECRE = Cells(TICKER, 9)
                   GREATDECRE = Cells(TICKER, 11)
                   End If
                   
                   
                   
                 TICKER = TICKER + 1
                 i = i - 1
                 sumadestocks = 0
                 contadorayuda = 1
                 End If
                
                Next
           
           
           Next
    
    
       Range("P2") = TICKERINCRE
       Range("P3") = TICKERDECRE
       Range("P4") = TICKERGREAT
       Range("Q2") = GREATINCRE
       Range("Q2").NumberFormat = "0.00%"
       Range("Q3") = GREATDECRE
       Range("Q3").NumberFormat = "0.00%"
       Range("Q4") = GREATTOTAL
GREATINCRE = 0
GREATDECRE = 0
GREATTOTAL = 0
        
    Next



End Sub

