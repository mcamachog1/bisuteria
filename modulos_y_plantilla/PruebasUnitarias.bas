Attribute VB_Name = "PruebasUnitarias"
Sub Principal()

    IniciarAmbienteParaPruebas
    'ProbarCostoUnitario
    probarObtenerFactor

End Sub
Sub IniciarAmbienteParaPruebas()

    InicializarVariablesGlobales
    InicializarFactoresDeConversion
    AMBIENTE = "DESARROLLO"
 
End Sub

Sub probarObtenerFactor()

    Dim tipoDeMedida As String
    Dim medidaOrigen As String
    Dim medidaDestino As String
    Dim mensaje1 As String
    Dim mensaje2 As String
    Dim mensaje3 As String
    Dim salidaEsperada As String
    Dim salidaObtenida As Variant
    Dim validacion As String
    
    On Error GoTo ManejadorDeErrores
    
    'CASO 1: factor de minuto a hora
    'SALIDA ESPERADA: 0.017
    salidaEsperada = 0.017
    mensaje1 = "CASO 1: factor de minuto a hora"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "TIEMPO"
    medidaOrigen = "MINUTO"
    cantidadDeUso = 1    '12 MILIGRAMOS
    medidaDestino = "HORA"
    cantidadUnidadesCompradas = 1 '9 KILOGRAMOS
    precio = 1
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 1
    
    Exit Sub
    
ManejadorDeErrores:

    salidaObtenida = Err.Number
    mensaje3 = "SALIDA OBTENIDA: " & Err.Description & " " & Err.Number
    If UCase(salidaEsperada) = UCase(salidaObtenida) Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    Err.Clear
    IniciarAmbienteParaPruebas
    Resume Next
    
End Sub

Sub ProbarCostoUnitario()

    Dim tipoDeMedida As String
    Dim medidaOrigen As String
    Dim medidaDestino As String
    Dim precio As Variant
    Dim cantidadDeUso As Variant
    Dim cantidadUnidadesCompradas As Variant
    Dim mensaje1 As String
    Dim mensaje2 As String
    Dim mensaje3 As String
    Dim salidaEsperada As String
    Dim salidaObtenida As Variant
    Dim validacion As String
    
    On Error GoTo ManejadorDeErrores
    
    'CASO 1: miligramos a kilogramos
    'SALIDA ESPERADA: 0
    salidaEsperada = 0
    mensaje1 = "CASO 1: miligramos a kilogramos"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "PESO"
    medidaOrigen = "MILIGRAMO"
    cantidadDeUso = 12    '12 MILIGRAMOS
    medidaDestino = "KILOGRAMO"
    cantidadUnidadesCompradas = 9 '9 KILOGRAMOS
    precio = 8.5
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 1

    'CASO 2: Factor no calculable
    'SALIDA ESPERADA: 1133333.333
    salidaEsperada = 1133333.333
    mensaje1 = "CASO 2: kilogramos a centigramos"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "PESO"
    medidaOrigen = "KILOGRAMO"
    cantidadDeUso = 12
    medidaDestino = "CENTIGRAMO"
    cantidadUnidadesCompradas = 9
    precio = 8.5
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 2

    'CASO 3: gramos a decigramos
    'SALIDA ESPERADA: 200
    salidaEsperada = 200
    mensaje1 = "CASO 3: gramos a decigramos"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "PESO"
    medidaOrigen = "GRAMO"
    cantidadDeUso = 10
    medidaDestino = "DECIGRAMO"
    cantidadUnidadesCompradas = 2
    precio = 4
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 3

    'CASO 4: metros a centimetros
    'SALIDA ESPERADA: 5000
    salidaEsperada = 5000
    mensaje1 = "CASO 4: metros a centimetros"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "LONGITUD"
    medidaOrigen = "METRO"
    cantidadDeUso = 5
    medidaDestino = "CENTÍMETRO"
    cantidadUnidadesCompradas = 1
    precio = 10
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 4

    'CASO 5: centimetros a metros
    'SALIDA ESPERADA: 0.113
    salidaEsperada = 0.113
    mensaje1 = "CASO 5: centimetros a metros"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "LONGITUD"
    medidaOrigen = "CENTÍMETRO"
    cantidadDeUso = 12
    medidaDestino = "METRO"
    cantidadUnidadesCompradas = 9
    precio = 8.5
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 5

    'CASO 6: minuto a semana
    'SALIDA ESPERADA: 0.089
    salidaEsperada = 0.089
    mensaje1 = "CASO 6: minuto a semana"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "TIEMPO"
    medidaOrigen = "MINUTO"
    cantidadDeUso = 45
    medidaDestino = "SEMANA"
    cantidadUnidadesCompradas = 1
    precio = 20
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 6


    'CASO 7: uso de porcentaje
    'SALIDA ESPERADA: 1
    salidaEsperada = 1
    mensaje1 = "CASO 7: uso de porcentaje"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "PORCENTAJE"
    medidaOrigen = "%"
    cantidadDeUso = 5
    medidaDestino = "%"
    cantidadUnidadesCompradas = 100 / 100
    precio = 20
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 7
    
    'CASO 8: medida que no existe
    'SALIDA ESPERADA: Error personalizado 1001
    salidaEsperada = "-2147220503"
    mensaje1 = "CASO 8: medida que no existe"
    mensaje2 = "SALIDA ESPERADA: En MATRIZ_FACTORES no se encuentra la medida de origen, de destino o ninguna de las dos. " & salidaEsperada
    tipoDeMedida = "PORCENTAJE"
    medidaOrigen = "PORCENTAJE"
    cantidadDeUso = 5
    medidaDestino = "PORCENTAJE"
    cantidadUnidadesCompradas = 100 / 100
    precio = 20
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    'FIN CASO 8
    
    'CASO 9: minuto a mes
    'SALIDA ESPERADA: 0.021
    salidaEsperada = 0.021
    mensaje1 = "CASO 9: minuto a mes"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "TIEMPO"
    medidaOrigen = "MINUTO"
    cantidadDeUso = 45
    medidaDestino = "MES"
    cantidadUnidadesCompradas = 1
    precio = 20
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 9
    
    'CASO 10: mes a hora
    'SALIDA ESPERADA: 0.005
    salidaEsperada = 0.005
    mensaje1 = "CASO 10: mes a hora"
    mensaje2 = "SALIDA ESPERADA: " & salidaEsperada
    tipoDeMedida = "TIEMPO"
    medidaOrigen = "HORA"
    cantidadDeUso = 0.75
    medidaDestino = "MES"
    cantidadUnidadesCompradas = 1
    precio = 5
    salidaObtenida = CostoUnitario(tipoDeMedida, medidaDestino, precio, medidaOrigen, cantidadDeUso, cantidadUnidadesCompradas)
    mensaje3 = "SALIDA OBTENIDA: " & CStr(salidaObtenida)
    If salidaEsperada = salidaObtenida Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    'FIN CASO 10
    
    
    Exit Sub

ManejadorDeErrores:

    salidaObtenida = Err.Number
    mensaje3 = "SALIDA OBTENIDA: " & Err.Description & " " & Err.Number
    If UCase(salidaEsperada) = UCase(salidaObtenida) Then
        validacion = "VALIDACION: TEST OK"
    Else
        validacion = "VALIDACION: TEST FAILED"
    End If
    Debug.Print mensaje1 & vbNewLine & mensaje2 & vbNewLine & mensaje3 & vbNewLine & validacion
    Err.Clear
    IniciarAmbienteParaPruebas
    Resume Next
     
End Sub


