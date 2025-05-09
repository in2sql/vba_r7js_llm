Attribute VB_Name = "Módulo1"
Option Explicit

' Función para enviar datos a Firebase al sistema de comidas
'****03/09/2024   SealtoSoft******
Sub EnviarDatosAFirebase()

    ' URLs de base de datos Firebase
    Dim URLD1, URLD2, URLD3, URLD4, URLD5, URLD6, URLD7, URLSemana, URLPedidos As String
    URLD1 = "https://foodalbaugh-default-rtdb.firebaseio.com/Dia1.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    URLD2 = "https://foodalbaugh-default-rtdb.firebaseio.com/Dia2.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    URLD3 = "https://foodalbaugh-default-rtdb.firebaseio.com/Dia3.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    URLD4 = "https://foodalbaugh-default-rtdb.firebaseio.com/Dia4.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    URLD5 = "https://foodalbaugh-default-rtdb.firebaseio.com/Dia5.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    URLD6 = "https://foodalbaugh-default-rtdb.firebaseio.com/Dia6.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    URLD7 = "https://foodalbaugh-default-rtdb.firebaseio.com/Dia7.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    
    URLSemana = "https://foodalbaugh-default-rtdb.firebaseio.com/Semana.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    URLPedidos = "https://foodalbaugh-default-rtdb.firebaseio.com/Pedidos.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""

    ' se definen variables de envio
    Dim D1O1, D1O2, D1O3 As String
    Dim D2O1, D2O2, D2O3 As String
    Dim D3O1, D3O2, D3O3 As String
    Dim D4O1, D4O2, D4O3 As String
    Dim D5O1, D5O2, D5O3 As String
    Dim D6O1, D6O2, D6O3 As String
    Dim D7O1, D7O2, D7O3 As String
    Dim Inicio, Final As String
    
    
    
    'se toman los datos de la celda y se guardan en variables
    D1O1 = Range("B4").Value
    D1O2 = Range("C4").Value
    D1O3 = Range("D4").Value
    
    D2O1 = Range("B5").Value
    D2O2 = Range("C5").Value
    D2O3 = Range("D5").Value
    
    D3O1 = Range("B6").Value
    D3O2 = Range("C6").Value
    D3O3 = Range("D6").Value
    
    D4O1 = Range("B7").Value
    D4O2 = Range("C7").Value
    D4O3 = Range("D7").Value
    
    D5O1 = Range("B8").Value
    D5O2 = Range("C8").Value
    D5O3 = Range("D8").Value
    
    D6O1 = Range("B9").Value
    D6O2 = Range("C9").Value
    D6O3 = Range("D9").Value
    
    D7O1 = Range("B10").Value
    D7O2 = Range("C10").Value
    D7O3 = Range("D10").Value
    
    Inicio = Range("B1").Value
    Final = Range("D1").Value
    
    
    Dim data1, data2, data3, data4, data5, data6, data7, semana, pedidos As String

    ' Crear los JSON para enviar a Firebase
    data1 = "{""op1"": """ & D1O1 & """, ""op2"": """ & D1O2 & """, ""op3"": """ & D1O3 & """}"
    data2 = "{""op1"": """ & D2O1 & """, ""op2"": """ & D2O2 & """, ""op3"": """ & D2O3 & """}"
    data3 = "{""op1"": """ & D3O1 & """, ""op2"": """ & D3O2 & """, ""op3"": """ & D3O3 & """}"
    data4 = "{""op1"": """ & D4O1 & """, ""op2"": """ & D4O2 & """, ""op3"": """ & D4O3 & """}"
    data5 = "{""op1"": """ & D5O1 & """, ""op2"": """ & D5O2 & """, ""op3"": """ & D5O3 & """}"
    data6 = "{""op1"": """ & D6O1 & """, ""op2"": """ & D6O2 & """, ""op3"": """ & D6O3 & """}"
    data7 = "{""op1"": """ & D7O1 & """, ""op2"": """ & D7O2 & """, ""op3"": """ & D7O3 & """}"
    semana = "{""Inicio"": """ & Inicio & """, ""Final"": """ & Final & """}"
    pedidos = "{""Semana"": """"}"
    

    ' Crear una solicitud HTTP
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Enviar las solicitudes PUT a Firebase
    With http
        .Open "PUT", URLD1, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send data1
    End With
    
    With http
        .Open "PUT", URLD2, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send data2
    End With

    With http
        .Open "PUT", URLD3, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send data3
    End With
    
    With http
        .Open "PUT", URLD4, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send data4
    End With
    
    With http
        .Open "PUT", URLD5, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send data5
    End With

    With http
        .Open "PUT", URLD6, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send data6
    End With

    With http
        .Open "PUT", URLD7, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send data7
    End With
    
     With http
        .Open "PUT", URLSemana, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send semana
    End With
    With http
        .Open "PUT", URLPedidos, False ' Puedes usar PUT si quieres actualizar un registro existente
        .setRequestHeader "Content-Type", "application/json"
        .send pedidos
    End With



    ' Comprobar la respuesta
    If http.Status = 200 Then
        MsgBox "Datos enviados correctamente."
    Else
        MsgBox "Error al enviar los datos: " & http.Status & " - " & http.statusText
    End If

End Sub
Sub AltaUsuarios()
    '03/09/2024
    'Macro para dar de alta en el sistema de pedido de comidas de android usando una base de datos firebase

    
    'Declaracion de variables necesarias
    Dim valor As String
    Dim cont As Integer
    Dim URLUsuario As String
    Dim data, data2 As String
    
    Dim Nombre, Pass As String
    Dim DNI As Long
    Dim DNI2 As String
    
    
    'se define la URL de la base de datos firebase
    URLUsuario = "https://foodalbaugh-default-rtdb.firebaseio.com/Usuario.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"""
    
    'se define el pass por defecto que es 1234 y cont es la variable que especifica en numero fila en la que va empezar a buscar datos el sistema
    Pass = "1234"
    cont = 3
    valor = Cells(cont, 1)
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    'se repite ciclo hasta que encuentra la celda sin ningun valor
    Do While valor <> ""
        DNI = Cells(cont, 1).Value
        DNI2 = Cells(cont, 1).Value
        Nombre = Cells(cont, 2).Value
        
        'es la Url que contiene el DNI para cargar los daros
        URLUsuario = "https://foodalbaugh-default-rtdb.firebaseio.com/Usuario/" + DNI2 + ".json"
        
        'con data se arma la estructura json y se guarda en una varibale
        data = "{""DNI"": " & DNI & ", ""nombre"": """ & Nombre & """, ""pass"": """ & Pass & """}"
        
        
        'con los datos en json y la Url se cargan los datos a firebase
        With http
            .Open "PUT", URLUsuario, False
            .setRequestHeader "Content-Type", "application/json"
            .send data
        End With
        
        cont = cont + 1
        valor = Cells(cont, 1)
    Loop
    
      ' Comprobar la respuesta
    If http.Status = 200 Then
        MsgBox "Datos enviados correctamente."
    Else
        MsgBox "Error al enviar los datos: " & http.Status & " - " & http.statusText
    End If
    
    
End Sub

Sub BlanquearClave()

    
    '03/09/2024
    'Macro para blanquear claves, en el sistema de pedido de comidas de android usando una base de datos firebase, lo que hace es poner el pass nuevamente en 1234

    Dim http As Object
    Dim url, data As String
    Dim respuesta As String
    Dim num As String
    
    num = InputBox("Ingrese el DNI de usuario a blanquear", "Blanquear Usuario")
    If num = "" Then Exit Sub

    ' URL de tu base de datos en Firebase
    url = "https://foodalbaugh-default-rtdb.firebaseio.com/Usuario/" + num + ".json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"

    ' Crear una instancia de WinHttp
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Configurar la solicitud
    http.Open "GET", url, False
    http.send

    ' Obtener la respuesta
    respuesta = http.responseText
    

    ' Mostrar la respuesta en la ventana de depuración (puedes ajustarlo según tus necesidades)
    If respuesta = "null" Then
        MsgBox "El usuario no esta cargado en la base de datos"
        Exit Sub
    End If
    
    Debug.Print respuesta

    ' Liberar el objeto
    Set http = Nothing
    
    ' Crear una solicitud HTTP
    ' Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Enviar la solicitud GET a Firebase
    With http
        .Open "GET", url, False
        .send
    End With

    ' Comprobar la respuesta
    If http.Status = 200 Then
        ' Si la respuesta es correcta, parsear el JSON
        Dim jsonResponse As String
        jsonResponse = http.responseText
        
        ' Parsear el JSON (utilizando Microsoft Scripting Runtime para trabajar con JSON)
        Dim jsonObject As Object
        Set jsonObject = JsonConverter.ParseJson(jsonResponse)
        
        ' Volcar los datos en la hoja de Excel
        Dim i As Integer
        i = 7
        Dim item As Variant
        
        For Each item In jsonObject
            ' Volcar los datos en las celdas A y B
            Debug.Print item
            Debug.Print jsonObject(item)
            Cells(i, 7).Value = jsonObject(item)
            i = i + 1
        Next item
        Dim DNI As Long
        Dim Nombre As String
        Dim Pass As String
        
        Pass = "1234"
        DNI = Cells(7, 7).Value
        Nombre = Cells(8, 7).Value
        
        data = "{""DNI"": " & DNI & ", ""nombre"": """ & Nombre & """, ""pass"": """ & Pass & """}"
        
        With http
            .Open "PUT", url, False
            .setRequestHeader "Content-Type", "application/json"
            .send data
        End With
        Cells(7, 7).Value = ""
        Cells(8, 7).Value = ""
        Cells(9, 7).Value = ""
        
        
        MsgBox "Clave Banqueada correctamente."
    Else
        MsgBox "Error al leer los datos: " & http.Status & " - " & http.statusText
    End If
End Sub

Sub LeerDatosDeFirebase()
    Call Borrar
    Dim http As Object
    Dim url As String
    Dim respuesta As String

    ' URL de tu base de datos en Firebase
    url = "https://foodalbaugh-default-rtdb.firebaseio.com/Pedidos/semana.json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"

    ' Crear una instancia de WinHttp
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Configurar la solicitud
    http.Open "GET", url, False
    http.send

    ' Obtener la respuesta
    respuesta = http.responseText
    If respuesta = "null" Then
        MsgBox "No hay pedidos cargados"
        Exit Sub
    End If
    Dim jsonObject As Object
        Set jsonObject = JsonConverter.ParseJson(respuesta)
        
        ' Volcar los datos en la hoja de Excel
        Dim i As Integer
       
        Dim item As Variant
        Dim sem As String
        
        
        For Each item In jsonObject
            ' Volcar los datos en las celdas A y B
            sem = item
            
           
        Next item
        url = "https://foodalbaugh-default-rtdb.firebaseio.com/Pedidos/semana/" + sem + ".json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"
         
        http.Open "GET", url, False
        http.send
        ' Obtener la respuesta
        respuesta = http.responseText
        Set jsonObject = JsonConverter.ParseJson(respuesta)
        i = 2
       
        
        For Each item In jsonObject
           ' Volcar los datos en las celdas A y B
           Cells(i, 1).Value = item
           i = i + 1
        Next item
        
        Dim cont As Integer
        Dim Contenido As String
    
        cont = 2
        Do While Cells(cont, 1).Value <> ""
            Contenido = Cells(cont, 1).Value
            url = "https://foodalbaugh-default-rtdb.firebaseio.com/Pedidos/semana/" + sem + "/" + Contenido + ".json?auth=AIzaSyDa-SN2dnwbY8PYvV-CXiex4g7yqNDi4Kc"
            http.Open "GET", url, False
            http.send
            ' Obtener la respuesta
            respuesta = http.responseText
            Set jsonObject = JsonConverter.ParseJson(respuesta)
            i = 2
            Dim colum As Integer
            Dim Obj2 As Object
            
            For Each item In jsonObject
               ' Volcar los datos en las celdas A y B
               colum = CInt(Mid(item, 1, 1))
               Cells(1, colum + 1).Value = item
               Set Obj2 = jsonObject(item)
               Cells(cont, colum + 1) = Obj2("Opcion")
               Debug.Print Obj2("Opcion")
               i = i + 1
            Next item
        
            cont = cont + 1
        Loop
        Dim rango As String
        rango = "A1:H" + Trim(Str(cont))
        
        
         'se da estilo a la celda para poder aplicar contadores y filtros
        Range(rango).Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$H$" + Trim(Str(cont - 1))), , xlYes).Name = _
        "Tabla7"
   
    ActiveSheet.ListObjects("Tabla7").TableStyle = "TableStyleMedium18"
    ActiveSheet.ListObjects("Tabla7").ShowTotals = True
    
    If CB1 <> "" Then
        Range("Tabla7[[#Totals],[" + CB1 + "]]").Select
        ActiveSheet.ListObjects("Tabla7").ListColumns(CB1). _
            TotalsCalculation = xlTotalsCalculationCount
    End If
    If CC1 <> "" Then
        Range("Tabla7[[#Totals],[" + CC1 + "]]").Select
        ActiveSheet.ListObjects("Tabla7").ListColumns(CC1). _
            TotalsCalculation = xlTotalsCalculationCount
    End If
    
    If CD1 <> "" Then
        Range("Tabla7[[#Totals],[" + CD1 + "]]").Select
        ActiveSheet.ListObjects("Tabla7").ListColumns(CD1). _
            TotalsCalculation = xlTotalsCalculationCount
    End If
    If CE1 <> "" Then
        Range("Tabla7[[#Totals],[" + CE1 + "]]").Select
        ActiveSheet.ListObjects("Tabla7").ListColumns(CE1). _
            TotalsCalculation = xlTotalsCalculationCount
    End If
    If CF1 <> "" Then
        Range("Tabla7[[#Totals],[" + CF1 + "]]").Select
        ActiveSheet.ListObjects("Tabla7").ListColumns(CF1). _
            TotalsCalculation = xlTotalsCalculationCount
    End If
    If CG1 <> "" Then
        Range("Tabla7[[#Totals],[" + CG1 + "]]").Select
        ActiveSheet.ListObjects("Tabla7").ListColumns(CG1). _
            TotalsCalculation = xlTotalsCalculationCount
    End If
    If CH1 <> "" Then
        Range("Tabla7[[#Totals],[" + CH1 + "]]").Select
        ActiveSheet.ListObjects("Tabla7").ListColumns(CH1). _
            TotalsCalculation = xlTotalsCalculationCount
    End If
  
        
    Set http = Nothing
End Sub
Sub Borrar()
'
' Borrar Macro
'

'
    Range("A1:H112").Select
    Selection.ClearContents
    Range("A1").Select
End Sub

