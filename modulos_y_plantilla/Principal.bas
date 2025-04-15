Attribute VB_Name = "Principal"
Public columnaNombreMaterial As String
Public columnaCostoMaterial As String
Public columnaNombreServicio As String
Public columnaCostoServicio As String
Public columnaNombreMdo As String
Public columnaCostoMdo As String
Public MATRIZ_FACTORES(1 To 65, 1 To 4) As Variant
Public parametros(1 To 10) As Variant
Public AMBIENTE As String
Public Const vbObjectError = -2147221504 ' Constante base para errores definidos por el usuario
Sub InicializarVariablesGlobales()
    columnaNombreMaterial = "D"
    columnaCostoMaterial = "E"
    columnaNombreServicio = "G"
    columnaCostoServicio = "H"
    columnaNombreMdo = "J"
    columnaCostoMdo = "K"
   
    AMBIENTE = "PRODUCCION"
    'AMBIENTE = "DESARROLLO"

       
End Sub
Sub GenerarErrorMedidaNoEncontrada(modulo As String, subrutina As String)
    'Const vbObjectError = -2147221504 ' Constante base para errores definidos por el usuario

    Dim numeroError As Long
    Dim origenError As String
    Dim descripcionError As String

    ' Define los detalles de tu error personalizado
    numeroError = vbObjectError + 1001 ' Usa una base y añade un número único
    origenError = modulo & "." & subrutina ' Especifica el origen del error
    descripcionError = "En MATRIZ_FACTORES no se encuentra la medida de origen, de destino o ninguna de las dos."

    ' Genera el error
    Err.Raise numeroError, origenError, descripcionError
End Sub
Sub prueba()
  GenerarErrorMedidaNoEncontrada "Principal", "prueba"

End Sub
'Recorre la matriz de factores y retorna el factor necesario para hacer el calculo
Function ObtenerFactor(tipoDeMedida As String, medidaOrigen As String, medidaDestino As String) As Variant

    Dim fila As Integer
    Dim factor As Variant
    
    If AMBIENTE = "DESARROLLO" Then
        On Error GoTo 0
    Else
        On Error GoTo ManejarError
    End If
    
        ' Recorrer el arreglo y comparar los valores
        For fila = 1 To UBound(MATRIZ_FACTORES)
            If MATRIZ_FACTORES(fila, 1) = LCase(tipoDeMedida) And MATRIZ_FACTORES(fila, 2) = LCase(medidaOrigen) And MATRIZ_FACTORES(fila, 3) = LCase(medidaDestino) Then
                ' Los 3 valores coinciden
                factor = CDec(MATRIZ_FACTORES(fila, 4))
                Exit For
            ElseIf MATRIZ_FACTORES(fila, 1) = LCase(tipoDeMedida) And MATRIZ_FACTORES(fila, 2) = LCase(medidaDestino) And MATRIZ_FACTORES(fila, 3) = LCase(medidaOrigen) Then
                factor = CDec(1 / MATRIZ_FACTORES(fila, 4))
                Exit For
            End If
        Next fila
        
    If IsEmpty(factor) Then
        GenerarErrorMedidaNoEncontrada "Principal", "ObtenerFactor"
        Exit Function
    End If
    If factor = 0 And factor <> "" Then
        Err.Raise 11, "CostosUnitarios", "División por cero"
    End If
    ObtenerFactor = factor
    Exit Function
ManejarError:
    ' Código para manejar el error

    If Err.Number = 11 Then
        MsgBox "La conversión de unidades no es razonable (Ejemplo: Kilometros a centímetros), escoja unidades más razonables", vbInformation, "Cálculo incompleto"
    End If
    ManejadorError ("ObtenerFactor")
    On Error GoTo 0
End Function
Sub ProvocarError()

    Dim cociente As Integer
    
    cociente = 1 / 2
    cociente = 1 / 0

End Sub
Sub InicializarFactoresDeConversion()
    FACTOR_TIEMPO_MIN_A_SEMANA = CDec(1) / CDec(10080)
    
    'Orden de la asignación:
    
    ',1 Tipo de medida
    ',2 Medida origen
    ',3 Medida destino
    ',4 Factor
    
    '------ INICIO MEDIDAS DE TIEMPO ------
    
    'MINUTO
    MATRIZ_FACTORES(1, 1) = "tiempo"
    MATRIZ_FACTORES(1, 2) = "minuto"
    MATRIZ_FACTORES(1, 3) = "segundo"
    MATRIZ_FACTORES(1, 4) = CDec(60)
    
    MATRIZ_FACTORES(2, 1) = "tiempo"
    MATRIZ_FACTORES(2, 2) = "minuto"
    MATRIZ_FACTORES(2, 3) = "minuto"
    MATRIZ_FACTORES(2, 4) = CDec(1)
    
    MATRIZ_FACTORES(3, 1) = "tiempo"
    MATRIZ_FACTORES(3, 2) = "minuto"
    MATRIZ_FACTORES(3, 3) = "hora"
    MATRIZ_FACTORES(3, 4) = CDec(1 / 60)
        
    MATRIZ_FACTORES(4, 1) = "tiempo"
    MATRIZ_FACTORES(4, 2) = "minuto"
    MATRIZ_FACTORES(4, 3) = "día"
    MATRIZ_FACTORES(4, 4) = CDec(1 / (60 * 24))
    
    MATRIZ_FACTORES(5, 1) = "tiempo"
    MATRIZ_FACTORES(5, 2) = "minuto"
    MATRIZ_FACTORES(5, 3) = "semana"
    MATRIZ_FACTORES(5, 4) = CDec(1 / (60 * 24 * 7))
    
    MATRIZ_FACTORES(6, 1) = "tiempo"
    MATRIZ_FACTORES(6, 2) = "minuto"
    MATRIZ_FACTORES(6, 3) = "mes"
    MATRIZ_FACTORES(6, 4) = 0 'Desbordamiento de memoria CDec(1 / (60 * 24 * 30))
    
    MATRIZ_FACTORES(7, 1) = "tiempo"
    MATRIZ_FACTORES(7, 2) = "minuto"
    MATRIZ_FACTORES(7, 3) = "año"
    MATRIZ_FACTORES(7, 4) = 0 'Desbordamiento de memoria CDec(1 / (60 * 24 * 365))
    
    'HORA
    MATRIZ_FACTORES(8, 1) = "tiempo"
    MATRIZ_FACTORES(8, 2) = "hora"
    MATRIZ_FACTORES(8, 3) = "segundo"
    MATRIZ_FACTORES(8, 4) = CDec(3600)
    
    'hora -> minuto está en inverso de minuto -> hora
    
    MATRIZ_FACTORES(9, 1) = "tiempo"
    MATRIZ_FACTORES(9, 2) = "hora"
    MATRIZ_FACTORES(9, 3) = "hora"
    MATRIZ_FACTORES(9, 4) = CDec(1)
        
    MATRIZ_FACTORES(10, 1) = "tiempo"
    MATRIZ_FACTORES(10, 2) = "hora"
    MATRIZ_FACTORES(10, 3) = "día"
    MATRIZ_FACTORES(10, 4) = CDec(1 / 24)
    
    MATRIZ_FACTORES(11, 1) = "tiempo"
    MATRIZ_FACTORES(11, 2) = "hora"
    MATRIZ_FACTORES(11, 3) = "semana"
    MATRIZ_FACTORES(11, 4) = CDec(1 / (24 * 7))
    
    MATRIZ_FACTORES(12, 1) = "tiempo"
    MATRIZ_FACTORES(12, 2) = "hora"
    MATRIZ_FACTORES(12, 3) = "mes"
    MATRIZ_FACTORES(12, 4) = CDec(1 / (24 * 30))
    
    MATRIZ_FACTORES(13, 1) = "tiempo"
    MATRIZ_FACTORES(13, 2) = "hora"
    MATRIZ_FACTORES(13, 3) = "año"
    MATRIZ_FACTORES(13, 4) = CDec(1 / (24 * 365))
    
    'DÍA
    MATRIZ_FACTORES(14, 1) = "tiempo"
    MATRIZ_FACTORES(14, 2) = "día"
    MATRIZ_FACTORES(14, 3) = "segundo"
    MATRIZ_FACTORES(14, 4) = 0 'Desbordamiento de memoria CDec(3600 * 24) combinación no factible
    
    'día -> minuto está en inverso de minuto -> día
    'día -> hora está en inverso de hora -> día
   
    MATRIZ_FACTORES(15, 1) = "tiempo"
    MATRIZ_FACTORES(15, 2) = "día"
    MATRIZ_FACTORES(15, 3) = "día"
    MATRIZ_FACTORES(15, 4) = CDec(1)
    
    MATRIZ_FACTORES(16, 1) = "tiempo"
    MATRIZ_FACTORES(16, 2) = "día"
    MATRIZ_FACTORES(16, 3) = "semana"
    MATRIZ_FACTORES(16, 4) = CDec(1 / 7)
    
    MATRIZ_FACTORES(17, 1) = "tiempo"
    MATRIZ_FACTORES(17, 2) = "día"
    MATRIZ_FACTORES(17, 3) = "mes"
    MATRIZ_FACTORES(17, 4) = CDec(1 / 30)
    
    MATRIZ_FACTORES(18, 1) = "tiempo"
    MATRIZ_FACTORES(18, 2) = "día"
    MATRIZ_FACTORES(18, 3) = "año"
    MATRIZ_FACTORES(18, 4) = CDec(1 / 365)
    
    'SEMANA
    MATRIZ_FACTORES(19, 1) = "tiempo"
    MATRIZ_FACTORES(19, 2) = "semana"
    MATRIZ_FACTORES(19, 3) = "segundo"
    MATRIZ_FACTORES(19, 4) = 0  'Desbordamiento de memoria CDec(3600 * 24 * 7) combinación no factible
    
    'semana -> minuto está en inverso de minuto -> semana
    'semana -> hora está en inverso de hora -> semana
    'semana -> día está en inverso de día -> semana
   
    MATRIZ_FACTORES(20, 1) = "tiempo"
    MATRIZ_FACTORES(20, 2) = "semana"
    MATRIZ_FACTORES(20, 3) = "semana"
    MATRIZ_FACTORES(20, 4) = CDec(1)
    
    MATRIZ_FACTORES(21, 1) = "tiempo"
    MATRIZ_FACTORES(21, 2) = "semana"
    MATRIZ_FACTORES(21, 3) = "mes"
    MATRIZ_FACTORES(21, 4) = CDec(7 / 30)
    
    MATRIZ_FACTORES(22, 1) = "tiempo"
    MATRIZ_FACTORES(22, 2) = "semana"
    MATRIZ_FACTORES(22, 3) = "año"
    MATRIZ_FACTORES(22, 4) = CDec(7 / 365)
    
    'MES
    MATRIZ_FACTORES(23, 1) = "tiempo"
    MATRIZ_FACTORES(23, 2) = "mes"
    MATRIZ_FACTORES(23, 3) = "segundo"
    MATRIZ_FACTORES(23, 4) = 0  'Desbordamiento de memoria CDec(3600 * 24 * 30) combinación no factible
    
    'mes -> minuto está en inverso de minuto -> mes
    'mes -> hora está en inverso de hora -> mes
    'mes -> día está en inverso de día -> mes
    'mes -> semana está en inverso de día -> mes
   
    MATRIZ_FACTORES(24, 1) = "tiempo"
    MATRIZ_FACTORES(24, 2) = "mes"
    MATRIZ_FACTORES(24, 3) = "mes"
    MATRIZ_FACTORES(24, 4) = CDec(1)
    
    MATRIZ_FACTORES(25, 1) = "tiempo"
    MATRIZ_FACTORES(25, 2) = "mes"
    MATRIZ_FACTORES(25, 3) = "año"
    MATRIZ_FACTORES(25, 4) = CDec(1 / 12)
    
    'AÑO
    MATRIZ_FACTORES(26, 1) = "tiempo"
    MATRIZ_FACTORES(26, 2) = "año"
    MATRIZ_FACTORES(26, 3) = "segundo"
    MATRIZ_FACTORES(26, 4) = 0 'Desbordamiento de memoria CDec(3600 * 24 * 365) combinación no factible
    
    'año -> minuto está en inverso de minuto -> año
    'año -> hora está en inverso de hora -> año
    'año -> día está en inverso de día -> año
    'año -> semana está en inverso de día -> año
    'año -> mes está en inverso de mes -> año
   
    MATRIZ_FACTORES(27, 1) = "tiempo"
    MATRIZ_FACTORES(27, 2) = "año"
    MATRIZ_FACTORES(27, 3) = "año"
    MATRIZ_FACTORES(27, 4) = CDec(1)

    'SEGUNDO
    MATRIZ_FACTORES(28, 1) = "tiempo"
    MATRIZ_FACTORES(28, 2) = "segundo"
    MATRIZ_FACTORES(28, 3) = "segundo"
    MATRIZ_FACTORES(28, 4) = CDec(1)
    
    'segundo -> minuto está en inverso de minuto -> segundo
    'segundo -> hora está en inverso de hora -> segundo
    'segundo -> día está en inverso de día -> segundo
    'segundo -> semana está en inverso de día -> segundo
    'segundo -> mes está en inverso de mes -> segundo
    'segundo -> año está en inverso de año -> segundo
   
    
    '------ INICIO MEDIDAS DE LONGITUD ------

    'CENTÍMETRO
    MATRIZ_FACTORES(29, 1) = "longitud"
    MATRIZ_FACTORES(29, 2) = "centímetro"
    MATRIZ_FACTORES(29, 3) = "milímetro"
    MATRIZ_FACTORES(29, 4) = CDec(10)

    MATRIZ_FACTORES(30, 1) = "longitud"
    MATRIZ_FACTORES(30, 2) = "centímetro"
    MATRIZ_FACTORES(30, 3) = "centímetro"
    MATRIZ_FACTORES(30, 4) = CDec(1)

    MATRIZ_FACTORES(31, 1) = "longitud"
    MATRIZ_FACTORES(31, 2) = "centímetro"
    MATRIZ_FACTORES(31, 3) = "metro"
    MATRIZ_FACTORES(31, 4) = CDec(1 / 100)

    'METRO
    MATRIZ_FACTORES(32, 1) = "longitud"
    MATRIZ_FACTORES(32, 2) = "metro"
    MATRIZ_FACTORES(32, 3) = "milímetro"
    MATRIZ_FACTORES(32, 4) = CDec(1000)

    'metro -> centímetro está en inverso de centímetro -> metro

    MATRIZ_FACTORES(33, 1) = "longitud"
    MATRIZ_FACTORES(33, 2) = "metro"
    MATRIZ_FACTORES(33, 3) = "metro"
    MATRIZ_FACTORES(33, 4) = CDec(1)

    'MILÍMETRO
    MATRIZ_FACTORES(34, 1) = "longitud"
    MATRIZ_FACTORES(34, 2) = "milímetro"
    MATRIZ_FACTORES(34, 3) = "milímetro"
    MATRIZ_FACTORES(34, 4) = CDec(1)
    'milímetro -> centímetro está en inverso de centímetro -> milímetro
    'milímetro -> metro está en inverso de milímetro -> metro

    '------ INICIO MEDIDAS DE UNIDAD ------
    
    'UNIDAD
    MATRIZ_FACTORES(35, 1) = "unidad"
    MATRIZ_FACTORES(35, 2) = "unidad"
    MATRIZ_FACTORES(35, 3) = "unidad"
    MATRIZ_FACTORES(35, 4) = 1
    
        
    '------ INICIO MEDIDAS DE CAPACIDAD ------
    'LITRO
    
    MATRIZ_FACTORES(36, 1) = "capacidad"
    MATRIZ_FACTORES(36, 2) = "litro"
    MATRIZ_FACTORES(36, 3) = "litro"
    MATRIZ_FACTORES(36, 4) = 1
    
    MATRIZ_FACTORES(37, 1) = "capacidad"
    MATRIZ_FACTORES(37, 2) = "litro"
    MATRIZ_FACTORES(37, 3) = "decilitro"
    MATRIZ_FACTORES(37, 4) = 10

    MATRIZ_FACTORES(38, 1) = "capacidad"
    MATRIZ_FACTORES(38, 2) = "litro"
    MATRIZ_FACTORES(38, 3) = "centilitro"
    MATRIZ_FACTORES(38, 4) = 100

    MATRIZ_FACTORES(39, 1) = "capacidad"
    MATRIZ_FACTORES(39, 2) = "litro"
    MATRIZ_FACTORES(39, 3) = "mililitro"
    MATRIZ_FACTORES(39, 4) = 1000
    
    
    'DECILITRO
    
    'decilitro -> litro está en inverso de litro -> decilitro

    MATRIZ_FACTORES(40, 1) = "capacidad"
    MATRIZ_FACTORES(40, 2) = "decilitro"
    MATRIZ_FACTORES(40, 3) = "decilitro"
    MATRIZ_FACTORES(40, 4) = 1

    MATRIZ_FACTORES(41, 1) = "capacidad"
    MATRIZ_FACTORES(41, 2) = "decilitro"
    MATRIZ_FACTORES(41, 3) = "centilitro"
    MATRIZ_FACTORES(41, 4) = 10

    MATRIZ_FACTORES(42, 1) = "capacidad"
    MATRIZ_FACTORES(42, 2) = "decilitro"
    MATRIZ_FACTORES(42, 3) = "mililitro"
    MATRIZ_FACTORES(42, 4) = 100
    
    'CENTILITRO
    
    'centilitro -> litro está en inverso de litro -> centilitro
    'centilitro -> decilitro está en inverso de decilitro -> centilitro
    
    MATRIZ_FACTORES(43, 1) = "capacidad"
    MATRIZ_FACTORES(43, 2) = "centilitro"
    MATRIZ_FACTORES(43, 3) = "centilitro"
    MATRIZ_FACTORES(43, 4) = 1

    MATRIZ_FACTORES(44, 1) = "capacidad"
    MATRIZ_FACTORES(44, 2) = "centilitro"
    MATRIZ_FACTORES(44, 3) = "mililitro"
    MATRIZ_FACTORES(44, 4) = 10
    
    'MILILITRO
    
    MATRIZ_FACTORES(45, 1) = "capacidad"
    MATRIZ_FACTORES(45, 2) = "mililitro"
    MATRIZ_FACTORES(45, 3) = "mililitro"
    MATRIZ_FACTORES(45, 4) = 1
        
        
    '------ INICIO MEDIDAS DE PORCENTAJE ------
    
    'PORCENTAJE

    MATRIZ_FACTORES(46, 1) = "porcentaje"
    MATRIZ_FACTORES(46, 2) = "%"
    MATRIZ_FACTORES(46, 3) = "%"
    MATRIZ_FACTORES(46, 4) = 1
    
    '------ INICIO MEDIDAS DE PESO ------
    
    'GRAMO
    
    MATRIZ_FACTORES(47, 1) = "peso"
    MATRIZ_FACTORES(47, 2) = "gramo"
    MATRIZ_FACTORES(47, 3) = "gramo"
    MATRIZ_FACTORES(47, 4) = 1
    
    MATRIZ_FACTORES(48, 1) = "peso"
    MATRIZ_FACTORES(48, 2) = "gramo"
    MATRIZ_FACTORES(48, 3) = "decigramo"
    MATRIZ_FACTORES(48, 4) = 10
            
    MATRIZ_FACTORES(49, 1) = "peso"
    MATRIZ_FACTORES(49, 2) = "gramo"
    MATRIZ_FACTORES(49, 3) = "centigramo"
    MATRIZ_FACTORES(49, 4) = 100
    
    MATRIZ_FACTORES(50, 1) = "peso"
    MATRIZ_FACTORES(50, 2) = "gramo"
    MATRIZ_FACTORES(50, 3) = "miligramo"
    MATRIZ_FACTORES(50, 4) = 1000
    
    'DECIGRAMO
    
    MATRIZ_FACTORES(51, 1) = "peso"
    MATRIZ_FACTORES(51, 2) = "decigramo"
    MATRIZ_FACTORES(51, 3) = "decigramo"
    MATRIZ_FACTORES(51, 4) = 1
    
    MATRIZ_FACTORES(52, 1) = "peso"
    MATRIZ_FACTORES(52, 2) = "decigramo"
    MATRIZ_FACTORES(52, 3) = "centigramo"
    MATRIZ_FACTORES(52, 4) = 10
    
    MATRIZ_FACTORES(53, 1) = "peso"
    MATRIZ_FACTORES(53, 2) = "decigramo"
    MATRIZ_FACTORES(53, 3) = "miligramo"
    MATRIZ_FACTORES(53, 4) = 100
    
    'CENTIGRAMO
    
    MATRIZ_FACTORES(54, 1) = "peso"
    MATRIZ_FACTORES(54, 2) = "centigramo"
    MATRIZ_FACTORES(54, 3) = "centigramo"
    MATRIZ_FACTORES(54, 4) = 1
    
    MATRIZ_FACTORES(55, 1) = "peso"
    MATRIZ_FACTORES(55, 2) = "centigramo"
    MATRIZ_FACTORES(55, 3) = "miligramo"
    MATRIZ_FACTORES(55, 4) = 10
    
    'MILIGRAMO
    MATRIZ_FACTORES(56, 1) = "peso"
    MATRIZ_FACTORES(56, 2) = "miligramo"
    MATRIZ_FACTORES(56, 3) = "miligramo"
    MATRIZ_FACTORES(56, 4) = 1
    
    'KILOGRAMO
    MATRIZ_FACTORES(57, 1) = "peso"
    MATRIZ_FACTORES(57, 2) = "kilogramo"
    MATRIZ_FACTORES(57, 3) = "kilogramo"
    MATRIZ_FACTORES(57, 4) = 1
    
    MATRIZ_FACTORES(58, 1) = "peso"
    MATRIZ_FACTORES(58, 2) = "kilogramo"
    MATRIZ_FACTORES(58, 3) = "gramo"
    MATRIZ_FACTORES(58, 4) = 1000
    
    MATRIZ_FACTORES(59, 1) = "peso"
    MATRIZ_FACTORES(59, 2) = "kilogramo"
    MATRIZ_FACTORES(59, 3) = "decigramo"
    MATRIZ_FACTORES(59, 4) = 10000
            
    MATRIZ_FACTORES(60, 1) = "peso"
    MATRIZ_FACTORES(60, 2) = "kilogramo"
    MATRIZ_FACTORES(60, 3) = "centigramo"
    MATRIZ_FACTORES(60, 4) = 0 '100000 Desbordamiento de memoria
    
    MATRIZ_FACTORES(61, 1) = "peso"
    MATRIZ_FACTORES(61, 2) = "kilogramo"
    MATRIZ_FACTORES(61, 3) = "miligramo"
    MATRIZ_FACTORES(61, 4) = 0 '1000000 Desbordamiento de memoria
        
End Sub

'Retorna un string con valores separados por comas correspondientes a los valores únicos de la columna pasada como parámetro
Function CrearListaDeValores(nombreTabla As String, nombreHoja As String, nombreColumna As String) As String
    
    Dim tabla As ListObject
    Dim columna As ListColumn
    Dim rangoDatos As Range
    Dim nombresUnicos As Object
    Dim diccionarioOrdenado As Object
    Dim nombre As Variant
    Dim lista As String
    Dim celda As Range

    On Error GoTo ManejarError
        ' Establecer referencias a la tabla y la columna
        Set tabla = ThisWorkbook.Sheets(nombreHoja).ListObjects(nombreTabla)
        Set columna = tabla.ListColumns(nombreColumna)
        Set rangoDatos = columna.DataBodyRange
    
        ' Verificar si el rango de datos está vacío
        If rangoDatos Is Nothing Then
            CrearListaDeValores = "" ' Retornar cadena vacía si no hay datos
            Exit Function
        End If
    
        ' Crear un objeto Dictionary para almacenar nombres únicos
        Set nombresUnicos = CreateObject("Scripting.Dictionary")
        ' Crear un objeto Dictionary para almacenar los nombres ordenados alfabeticamente
        Set diccionarioOrdenado = CreateObject("Scripting.Dictionary")
    
        ' Recorrer el rango de datos y agregar nombres únicos al Dictionary
        For Each celda In rangoDatos
            If Not nombresUnicos.Exists(UCase(celda.Value)) Then
                nombresUnicos.Add UCase(celda.Value), 1
            End If
        Next celda
    
        ' Ordenar el diccionario (asumiendo que tienes una función OrdenarDiccionarioPorClave)
        Set diccionarioOrdenado = OrdenarDiccionarioPorClave(nombresUnicos)
    
        ' Crear la cadena de la lista
        For Each nombre In diccionarioOrdenado.Keys
            lista = lista & UCase(nombre) & ","
        Next nombre
    
        ' Eliminar la coma final (solo si la lista no está vacía)
        If Len(lista) > 0 Then
            lista = Left(lista, Len(lista) - 1)
        End If
    
        CrearListaDeValores = lista
    Exit Function

ManejarError:
    ' Código para manejar el error
    ManejadorError ("CrearListaDeValores")
    On Error GoTo 0

End Function

' En la hoja "Formulario" se hacen las validaciones en las tablas: Materiales, Servicios y ManoDeObra
' de los valores permitidos
Sub CrearListaDesplegable(nombreTabla As String, nombreHoja As String, nombreColumna As String, Optional tipoDeCosto As String = "")

Dim celda As Range
Dim lista As String
Dim tabla As ListObject
Dim ultimaFila As Long
Dim i As Long
    
On Error GoTo ManejarError
    ' Crea lista de valores permitidos
    lista = CrearListaDeValores(nombreTabla, nombreHoja, nombreColumna)
    
    ' Estandariza tipoDeCosto
    If UCase(tipoDeCosto) = "MATERIALES" Then
        tipoDeCosto = "Materiales"
    ElseIf UCase(tipoDeCosto) = "SERVICIOS" Then
        tipoDeCosto = "Servicios"
    ElseIf UCase(tipoDeCosto) = "MANODEOBRA" Then
        tipoDeCosto = "ManoDeObra"
    End If
    
    
    If Len(tipoDeCosto) = 0 Then ' La lista de valores es para CATEGORIAS celda B7
        Set celda = ThisWorkbook.Sheets("Formulario").Range("B7")
        With celda.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=lista
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else ' Se llena la tabla correspondiente, Materiales, Servicios o Mano de Obra
        Set tabla = ThisWorkbook.Sheets("Formulario").ListObjects("TablaF_" & tipoDeCosto)
    
        ' Encontrar la última fila con datos en la tabla
        ultimaFila = tabla.DataBodyRange.Rows(tabla.DataBodyRange.Rows.Count).Row
    
        ' Recorrer la tabla desde la fila 1 hasta la última fila
        For i = 1 To ultimaFila - tabla.HeaderRowRange.Row
            Set celda = tabla.DataBodyRange.Cells(i, 1)
            ' Aplicar la validación de datos de lista desplegable
            With celda.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=lista
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        Next i
    
        ' Limpiar la variable de objeto
        Set tabla = Nothing
        Set celda = Nothing
    End If
Exit Sub
ManejarError:
    ' Código para manejar el error
    ManejadorError ("CrearListaDesplegable")
    On Error GoTo 0
    
End Sub
Sub LimpiarFormularios()

    Dim tbl As ListObject
    Dim ws As Worksheet
    Dim categoria As String
    
    'Hoja de cálculo donde están las tablas a borrar
    Set ws = ThisWorkbook.Sheets("Formulario")
    
    'Verifica si hay tablas en la hoja
    If ws.ListObjects.Count > 0 Then
        'Recorre todas las tablas de la hoja y borra el contenido
        For Each tbl In ws.ListObjects
            tbl.DataBodyRange.ClearContents
        Next tbl
    Else
        MsgBox "No hay tablas en la hoja '" & ws.Name & "'."
    End If
    
    'Borra el nombre, categoria y margen del producto
    ws.Range("B4").ClearContents
    ws.Range("B7").ClearContents
    ws.Range("B10").ClearContents
    ws.Range("B13").ClearContents
    
End Sub
Sub MostrarFormularioMateriales()
    Materiales.Show
End Sub
Sub MostrarFormularioServicios()
    Servicios.Show
End Sub
Sub MostrarFormularioServiciosEditar(nombre As String, celda As Range)



    Dim miTabla As ListObject
    Dim valoresUnicos As Variant

    ' Establecer la tabla (reemplaza "Tabla1" y "Hoja1" con los nombres correctos)
    Set miTabla = ThisWorkbook.Sheets("UNIDADES").ListObjects("MaestroUnidades")

    ' Obtener los valores únicos de la columna "Nombre" (reemplaza con el nombre de tu columna)
    valoresUnicos = ValoresUnicosColumnaTabla(miTabla, "TIPO DE MEDIDA CON EJEMPLO")

    ' Limpiar el ComboBox antes de agregar nuevos valores
    ServiciosEditar.ComboBoxTipoDeMedida.Clear

    ' Agregar los valores únicos al ComboBox
    If UBound(valoresUnicos) > 0 Then
        For i = 1 To UBound(valoresUnicos)
            ServiciosEditar.ComboBoxTipoDeMedida.AddItem valoresUnicos(i)
        Next i
    Else
        MsgBox "No se encontraron valores únicos o la columna no existe."
    End If
    
    ServiciosEditar.txtNombre = UCase(nombre)
    ServiciosEditar.label_row = celda.Row
    ServiciosEditar.label_column = celda.Column

    ServiciosEditar.Show
    
End Sub
Sub MostrarFormularioMaterialesEditar(nombre As String, celda As Range)

    Dim miTabla As ListObject
    Dim valoresUnicos As Variant

    Set miTabla = ThisWorkbook.Sheets("UNIDADES").ListObjects("MaestroUnidades")

    ' Obtener los valores únicos de la columna "Nombre" (reemplaza con el nombre de tu columna)
    valoresUnicos = ValoresUnicosColumnaTabla(miTabla, "TIPO DE MEDIDA CON EJEMPLO")

    ' Limpiar el ComboBox antes de agregar nuevos valores
    MaterialesEditar.ComboBoxTipoDeMedida.Clear

    ' Agregar los valores únicos al ComboBox
    If UBound(valoresUnicos) > 0 Then
        For i = 1 To UBound(valoresUnicos)
            MaterialesEditar.ComboBoxTipoDeMedida.AddItem valoresUnicos(i)
        Next i
    Else
        MsgBox "No se encontraron valores únicos o la columna no existe."
    End If
    
    MaterialesEditar.txtNombre = UCase(nombre)
    MaterialesEditar.label_row = celda.Row
    MaterialesEditar.label_column = celda.Column

    MaterialesEditar.Show

End Sub
Sub MostrarFormularioManoDeObraEditar(nombre As String, celda As Range)
    Dim miTabla As ListObject
    Dim valoresUnicos As Variant

    Set miTabla = ThisWorkbook.Sheets("UNIDADES").ListObjects("MaestroUnidades")

    ' Obtener los valores únicos de la columna "Nombre" (reemplaza con el nombre de tu columna)
    valoresUnicos = ValoresUnicosColumnaTabla(miTabla, "TIPO DE MEDIDA CON EJEMPLO")

    ' Limpiar el ComboBox antes de agregar nuevos valores
    ManoDeObraEditar.ComboBoxTipoDeMedida.Clear

    ' Agregar los valores únicos al ComboBox
    If UBound(valoresUnicos) > 0 Then
        For i = 1 To UBound(valoresUnicos)
            ManoDeObraEditar.ComboBoxTipoDeMedida.AddItem valoresUnicos(i)
        Next i
    Else
        MsgBox "No se encontraron valores únicos o la columna no existe."
    End If
    
    ManoDeObraEditar.txtNombre = UCase(nombre)
    ManoDeObraEditar.label_row = celda.Row
    ManoDeObraEditar.label_column = celda.Column

    ManoDeObraEditar.Show
End Sub
Sub MostrarFormularioManoDeObra()
    ManoDeObra.Show
End Sub
Sub MostrarFormularioCategorias()
    On Error GoTo ManejarError

    Categorias.Show
    'Err.Raise 1001, "MostrarFormularioCategorias", "Error programado"
    
ManejarError:
    ManejadorError ("MostrarFormularioCategorias")

End Sub

Sub InsertarAlFinal(nombreHoja As String, nombreTabla As String, valor As String)

    'Buscar la ultima linea de la tabla MaestroMateriales de la hoja MATERIALES
    
    Dim tabla As ListObject
    Dim ultimaFila As Long

    ' Establecer la tabla (reemplaza "MATERIALES" y "MaestroMateriales" si es necesario)
    Set tabla = ThisWorkbook.Sheets(nombreHoja).ListObjects(nombreTabla)

    ' Verificar si la tabla tiene datos
    If Not tabla.DataBodyRange Is Nothing Then
        ' Encontrar la última fila con datos
        ultimaFila = tabla.DataBodyRange.Rows(tabla.DataBodyRange.Rows.Count).Row

        ' Agregar una nueva fila al final de la tabla
        Set nuevaFila = tabla.ListRows.Add

        ' Mostrar un mensaje indicando que se agregó la nueva fila
        'MsgBox "Se ha agregado una nueva fila al final de la tabla.", vbInformation

    Else
        ' Si la tabla está vacía, agregar una nueva fila
        Set nuevaFila = tabla.ListRows.Add
        'MsgBox "La tabla estaba vacía. Se ha agregado una nueva fila.", vbInformation
    End If

    nuevaFila.Range.Cells(1, 1).Value = UCase(valor)
    ' Limpiar las variables de objeto
    Set tabla = Nothing
    Set nuevaFila = Nothing
End Sub


'Elimina un producto y sus datos de costo de la FactTable
Sub EliminarProducto()
    Dim producto As String
    Dim respuesta As Integer
    Dim tbl As ListObject
        
   ' Intenta obtener la tabla "FactTable"
    On Error Resume Next
      Set tbl = Hoja5.ListObjects("FactTable")
    On Error GoTo 0
    
    If tbl Is Nothing Then
      MsgBox "La tabla 'FactTable' no existe en la hoja 'Hoja5'."
      Exit Sub
    End If
        
    producto = Hoja3.Range("B4").Value
    ' Se confirma la eliminación y se ejecuta
    respuesta = MsgBox("¿Está seguro de que desea eliminar '" & UCase(producto) & "'", vbYesNo + vbQuestion, "Confirmar eliminación")
    If respuesta = vbYes Then
        EliminarFilasPorValorColumna producto, "nombreProducto", tbl
        AgregarListaDesplegableDesdeTabla
        LimpiarFormularios
        MsgBox "Producto '" & UCase(producto) & "' eliminado"
    End If
    ActualizarTablasDinamicas
    
End Sub
'Dado una categoria de costo y un producto, eliminar los elementos duplicados y dejar el primer valor encontrado en factTable
'Ejemplo: El producto 'COLLAR DE PERLAS' tiene materiales repetidos, dejo el primero y elimino los demas
Sub EliminarDuplicadosDeFactTable(valorNombreProducto As String, valorCategoriaDeCosto As String)
    Dim tabla As ListObject
    Dim rangoTabla As Range
    Dim columnaCategoriaDeCosto As Long
    Dim columnaNombreProducto As Long
    Dim columnaMaterialServicioManoDeObra As Long
    Dim celda As Range
    Dim filasVisible As Range
    Dim diccionario As Object
    Dim clave As Variant
    Dim filasAEliminar As Range

    If UCase(valorCategoriaDeCosto) = "MATERIAL" Then
        columnaMaterialServicioManoDeObra = 4
    ElseIf UCase(valorCategoriaDeCosto) = "SERVICIO" Then
        columnaMaterialServicioManoDeObra = 10
    ElseIf UCase(valorCategoriaDeCosto) = "MANO DE OBRA" Then
        columnaMaterialServicioManoDeObra = 16
    ElseIf UCase(valorCategoriaDeCosto) = "MARGEN" Then
        columnaMaterialServicioManoDeObra = 22
    End If
    
    
    Set tabla = Hoja5.ListObjects("FactTable")
    
    ' Establecer el rango de la tabla
    Set rangoTabla = tabla.DataBodyRange
    
    columnaCategoriaDeCosto = 3
    columnaNombreProducto = 1
    
    rangoTabla.AutoFilter Field:=columnaNombreProducto, Criteria1:=UCase(valorNombreProducto)
    rangoTabla.AutoFilter Field:=columnaCategoriaDeCosto, Criteria1:=UCase(valorCategoriaDeCosto)
    
    ' Crear el diccionario
    Set diccionario = CreateObject("Scripting.Dictionary")
    
    ' Obtener solo las celdas visibles después del filtro
    On Error Resume Next ' Ignorar errores si no hay celdas visibles
    Set filasVisibles = rangoTabla.SpecialCells(xlCellTypeVisible).Rows
    On Error GoTo 0 ' Restablecer el manejo de errores

    ' Recorrer las filas visibles
    If Not filasVisibles Is Nothing Then
        For Each filaVisible In filasVisibles
            ' Obtener el valor de la columna que deseas verificar (ejemplo: columna 2)
            clave = filaVisible.Cells(1, columnaMaterialServicioManoDeObra).Value

            ' Verificar si la clave ya existe en el diccionario
            If diccionario.Exists(clave) Then
                ' Si la clave existe, agregar la fila a las filas a eliminar
                If filasAEliminar Is Nothing Then
                    Set filasAEliminar = filaVisible
                Else
                    Set filasAEliminar = Union(filasAEliminar, filaVisible)
                End If
            Else
                ' Si la clave no existe, agregarla al diccionario
                diccionario.Add clave, 1
            End If
        Next filaVisible
    Else
        MsgBox "No se encontraron filas visibles después del filtro."
    End If


    ' Desactivar las alertas de Excel
    Application.DisplayAlerts = False

    ' Eliminar las filas duplicadas
    If Not filasAEliminar Is Nothing Then
        filasAEliminar.Delete
    End If

    ' Reactivar las alertas de Excel
    Application.DisplayAlerts = True


    ' Eliminar el filtro
    tabla.Range.AutoFilter
    
      

End Sub
Sub ejemplo()
 EliminarDuplicadosDeFactTable "COLLAR DE PERLAS", "MARGEN"
End Sub
'Agrega un nuevo producto
Sub AgregarProducto()
  Dim tbl As ListObject
  Dim tblMateriales As ListObject
  Dim nuevaFila As ListRow

  'Si el producto no fue indicado salgo de la rutina
  If Len(Hoja3.Range("B4").Value) = 0 Then
    MsgBox "Escriba el nombre del producto en la celda B4", vbInformation
    Exit Sub
  End If
  
  'Si la categoría no fue indicada salgo de la rutina
  If Len(Hoja3.Range("B7").Value) = 0 Then
    MsgBox "Seleccione la categoría en la celda B7", vbInformation
    Exit Sub
  End If
  
  'Si el margen no fue indicado salgo de la rutina
  If Len(Hoja3.Range("B10").Value) = 0 Then
    MsgBox "Escriba el % de ganancia en la celda B10", vbInformation
    Exit Sub
  End If
 
  ' Intenta obtener la tabla "FactTable"
  On Error Resume Next
    Set tbl = Hoja5.ListObjects("FactTable")
  On Error GoTo 0

  If tbl Is Nothing Then
    MsgBox "La tabla 'FactTable' no existe en la hoja 'Hoja5'."
    Exit Sub
  End If
  
On Error GoTo ManejarError

  'Antes de agregar, elimino el producto (la subrutina maneja el caso de que el producto no exista)
  EliminarFilasPorValorColumna Hoja3.Range("B4").Value, "nombreProducto", tbl

  'Grabo los materiales
  Set tblMateriales = Hoja3.ListObjects("TablaF_Materiales")
  'Si no hay valor en la tabla salgo de la rutina
  If TablaVacia(tblMateriales) Then
    MsgBox "Escribe un material para poder guardar el producto"
    Exit Sub
  End If
  If Not tblMateriales.DataBodyRange Is Nothing Then
        ' Recorre cada fila de la tabla
        For Each fila In tblMateriales.ListRows
            'Si el material no es vacio se graba
             If Len(fila.Range.Cells(1, 1)) <> 0 Then
                ' Agrega una nueva fila al final de la tabla FactTable
                Set nuevaFila = tbl.ListRows.Add
                nuevaFila.Range(1, 1).Value = UCase(Hoja3.Range("B4")) 'Nombre producto
                nuevaFila.Range(1, 2).Value = UCase(Hoja3.Range("B7")) 'Categoria
                nuevaFila.Range(1, 3).Value = "Material" 'Tipo de Costo
                'Columna 4 de FactTable y Columna 1 de la tabla Materiales tiene el material
                nuevaFila.Range(1, 4).Value = UCase(fila.Range.Cells(1, 1))
                'Columna 9 de FacTable y Columna 2 de la tabla Materiales tiene el costo
                nuevaFila.Range(1, 9).Value = fila.Range.Cells(1, 2)
             End If
        Next fila
  Else
        MsgBox "La tabla '" & tblMateriales.Name & "' no tiene filas de datos."
  End If
 
  'Grabo los servicios
  Set tblServicios = Hoja3.ListObjects("TablaF_Servicios")
  If Not tblServicios.DataBodyRange Is Nothing Then
        For Each fila In tblServicios.ListRows
             If Len(fila.Range.Cells(1, 1)) <> 0 Then
                ' Agrega una nueva fila al final de la tabla FactTable
                Set nuevaFila = tbl.ListRows.Add
                nuevaFila.Range(1, 1).Value = UCase(Hoja3.Range("B4"))
                nuevaFila.Range(1, 2).Value = UCase(Hoja3.Range("B7"))
                nuevaFila.Range(1, 3).Value = "Servicio"
                'Columna 10 de FactTable y Columna 1 de la tabla Servicios tiene el servicio
                nuevaFila.Range(1, 10).Value = UCase(fila.Range.Cells(1, 1))
                'Columna 15 de FacTable y Columna 2 de la tabla Servicios tiene el costo
                nuevaFila.Range(1, 15).Value = fila.Range.Cells(1, 2)
             End If
        Next fila
  Else
        MsgBox "La tabla '" & tblServicios.Name & "' no tiene filas de datos."
  End If
 
 
  'Grabo la mano de obra
  Set tblManoDeObra = Hoja3.ListObjects("TablaF_ManoDeObra")
  If Not tblManoDeObra.DataBodyRange Is Nothing Then
        For Each fila In tblManoDeObra.ListRows
             If Len(fila.Range.Cells(1, 1)) <> 0 Then
                ' Agrega una nueva fila al final de la tabla FactTable
                Set nuevaFila = tbl.ListRows.Add
                nuevaFila.Range(1, 1).Value = UCase(Hoja3.Range("B4"))
                nuevaFila.Range(1, 2).Value = UCase(Hoja3.Range("B7"))
                nuevaFila.Range(1, 3).Value = "Mano de Obra"
                'Columna 16 de FactTable y Columna 1 de la tabla Mano De Obra tiene la descripcion
                nuevaFila.Range(1, 16).Value = UCase(fila.Range.Cells(1, 1))
                'Columna 21 de FacTable y Columna 2 de la tabla Mano De Obra tiene el costo
                nuevaFila.Range(1, 21).Value = fila.Range.Cells(1, 2)
             End If
        Next fila
  Else
        MsgBox "La tabla '" & tblManoDeObra.Name & "' no tiene filas de datos."
  End If
 
  'Grabo el margen de ganancia
  Set nuevaFila = tbl.ListRows.Add
  nuevaFila.Range(1, 1).Value = UCase(Hoja3.Range("B4"))
  nuevaFila.Range(1, 2).Value = UCase(Hoja3.Range("B7"))
  nuevaFila.Range(1, 3).Value = "Margen"
  'Columna 22 de FactTable y celda B10
  nuevaFila.Range(1, 22).Value = Hoja3.Range("B10")

  MsgBox "'" & Hoja3.Range("B4").Value & "' guardado(a) con éxito.", vbInformation, "Éxito"
  ActualizarTablasDinamicas
  AgregarListaDesplegableDesdeTabla

Exit Sub
ManejarError:
    ' Código para manejar el error
    ManejadorError ("AgregarProducto")
    On Error GoTo 0
  
End Sub
Function TablaVacia(tabla As ListObject) As Boolean
    If tabla.DataBodyRange Is Nothing Then
        TablaVacia = True
    ElseIf tabla.DataBodyRange.Rows.Count = 0 Then
        TablaVacia = True
    Else
        TablaVacia = False
    End If
    
    
End Function
' Dado un valor, una columna y un objeto tabla elimina las filas de la tabla que tienen ese valor en la columna
Sub EliminarFilasPorValorColumna(valor As String, nombreColumna As String, tabla As ListObject)
    Dim columna As Range
    Dim fila As Range
    Dim filasAEliminar As Range
    
    On Error GoTo ManejarError

    ' Encuentra la columna por nombre
    Set columna = tabla.HeaderRowRange.Find(nombreColumna, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica si se encontró la columna
    If columna Is Nothing Then
        MsgBox "No se encontró la columna '" & nombreColumna & "' en la tabla.", vbExclamation
        Exit Sub
    End If

    ' Recorre las filas de la tabla y marca las filas para eliminar
    For Each fila In tabla.DataBodyRange.Rows
        If UCase(fila.Cells(columna.Column).Value) = UCase(valor) Then
            If filasAEliminar Is Nothing Then
                Set filasAEliminar = fila
            Else
                Set filasAEliminar = Union(filasAEliminar, fila)
            End If
        End If
    Next fila

    ' Elimina las filas marcadas
    If Not filasAEliminar Is Nothing Then
        filasAEliminar.Delete
    Else
        'MsgBox "El producto '" & valor & "' es un producto nuevo", vbInformation
    End If
    
    Exit Sub
ManejarError:
    ManejadorError ("EliminarFilasPorValorColumna")
    On Error GoTo 0
    
End Sub

' Carga los datos del producto seleccionado en la celda B13
Sub EditarProducto(nombreProducto As String)

    Dim tabla As ListObject
    Dim columnaProducto, columnaCategoriaCostos As Range
    Dim i As Integer
    Dim miDiccionario As Object
    Set miDiccionario = CreateObject("Scripting.Dictionary")
    Dim clave As Variant
    Dim valor As Variant
    Dim elemento As String

    On Error GoTo ManejarError

    ' Establecer la referencia a la tabla
    Set tabla = ThisWorkbook.Sheets("FactTable").ListObjects("FactTable")

    ' Encuentra las columnas por nombre
    Set columnaProducto = tabla.HeaderRowRange.Find("nombreProducto", LookIn:=xlValues, LookAt:=xlWhole)
    Set columnaCategoriaCostos = tabla.HeaderRowRange.Find("categoriaCostos", LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica si se encontraron las columnas
    If columnaProducto Is Nothing Or columnaCategoriaCostos Is Nothing Then
        MsgBox "No se encontraron las columnas 'nombreProducto' o 'categoriaCostos'.", vbExclamation
    End If
    
    ' MATERIAL
    ' Aplica el filtro con ambos criterios
    tabla.Range.AutoFilter Field:=columnaProducto.Column, Criteria1:=nombreProducto
    tabla.Range.AutoFilter Field:=columnaCategoriaCostos.Column, Criteria1:="Material"
     
    ' Recorrer las filas visibles después del filtro y graba en un diccionario
    If Not tabla.DataBodyRange.SpecialCells(xlCellTypeVisible) Is Nothing Then
        miDiccionario.RemoveAll
        For Each fila In tabla.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows
            miDiccionario.Add fila.Cells(1, 4).Value, fila.Cells(1, 9).Value
        Next fila
    End If
    
    ' Valida que los materiales no se pasen de 12 porque si se pasan no caben en la tabla
    If miDiccionario.Count > 12 Then
        MsgBox "Error: Excede las 12 filas de la tabla materiales"
    Else
        ' Limpiar tabla Materiales
        For i = 4 To 14
            Hoja3.Range(columnaNombreMaterial & i).ClearContents
            Hoja3.Range(columnaCostoMaterial & i).ClearContents
        Next i
        ' Llenar tabla materiales
        i = 4
        For Each clave In miDiccionario.Keys
            Hoja3.Hyperlinks.Add Anchor:=Hoja3.Range(columnaNombreMaterial & i), Address:="", TextToDisplay:=UCase(clave)
            Hoja3.Range(columnaCostoMaterial & i).Value = miDiccionario.Item(clave)
            i = i + 1
        Next clave
    End If

    
    ' SERVICIO
    ' Aplica el 2do filtro con criterio categoriaDeCosto = "Servicio"
    tabla.Range.AutoFilter Field:=columnaCategoriaCostos.Column, Criteria1:="Servicio"
    ' Recorrer las filas visibles después del filtro y graba en un diccionario
    If Not tabla.DataBodyRange.SpecialCells(xlCellTypeVisible) Is Nothing Then
        miDiccionario.RemoveAll
        For Each fila In tabla.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows
            miDiccionario.Add fila.Cells(1, 10).Value, fila.Cells(1, 15).Value
        Next fila
    End If
    
    ' Valida que los servicios no se pasen de 12 porque si se pasan no caben en la tabla
    If miDiccionario.Count > 12 Then
        MsgBox "Error: Excede las 12 filas de la tabla servicios"
    Else
        ' Limpiar tabla Servicios
        For i = 4 To 14
            Hoja3.Range("G" & i).ClearContents
            Hoja3.Range("H" & i).ClearContents
        Next i
        
        ' Llenar tabla servicios
        i = 4
        For Each clave In miDiccionario.Keys
            Hoja3.Hyperlinks.Add Anchor:=Hoja3.Range(columnaNombreServicio & i), Address:="", TextToDisplay:=UCase(clave)
            Hoja3.Range(columnaCostoServicio & i).Value = miDiccionario.Item(clave)
            i = i + 1
        Next clave
    End If


    ' MANO DE OBRA
    ' Aplica el 2do filtro con criterio categoriaDeCosto = "Mano De Obra"
    tabla.Range.AutoFilter Field:=columnaCategoriaCostos.Column, Criteria1:="Mano de Obra"
    ' Recorrer las filas visibles después del filtro y graba en un diccionario
    If Not tabla.DataBodyRange.SpecialCells(xlCellTypeVisible) Is Nothing Then
        miDiccionario.RemoveAll
        For Each fila In tabla.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows
            miDiccionario.Add fila.Cells(1, 16).Value, fila.Cells(1, 21).Value
        Next fila
    End If
    
    ' Valida que no se pasen de 12 porque si se pasan no caben en la tabla
    If miDiccionario.Count > 12 Then
        MsgBox "Error: Excede las 12 filas de la tabla Mano de Obra"
    Else
        ' Limpiar tabla
        For i = 4 To 14
            Hoja3.Range(columnaNombreMdo & i).ClearContents
            Hoja3.Range(columnaCostoMdo & i).ClearContents
        Next i
        
        ' Llenar tabla Mano de Obra
        i = 4
        For Each clave In miDiccionario.Keys
            Hoja3.Hyperlinks.Add Anchor:=Hoja3.Range(columnaNombreMdo & i), Address:="", TextToDisplay:=UCase(clave)
            Hoja3.Range(columnaCostoMdo & i).Value = miDiccionario.Item(clave)
            i = i + 1
        Next clave
    End If

 
    ' MARGEN
    ' Aplica el 2do filtro con criterio categoriaDeCosto = "Margen"
    tabla.Range.AutoFilter Field:=columnaCategoriaCostos.Column, Criteria1:="Margen"
    
    ' Recorrer las filas visibles después del filtro y graba en un diccionario
    If Not tabla.DataBodyRange.SpecialCells(xlCellTypeVisible) Is Nothing Then
        miDiccionario.RemoveAll
        For Each fila In tabla.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows
            miDiccionario.Add fila.Cells(1, 22).Value, fila.Cells(1, 22).Value
        Next fila
    End If
    
    ' Seleccionar el valor del margen
    If Not tabla.DataBodyRange.SpecialCells(xlCellTypeVisible) Is Nothing Then
        For Each fila In tabla.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows
            Hoja3.Range("B10").Value = fila.Cells(1, 22).Value ' Porcentaje
            Hoja3.Range("B4").Value = UCase(fila.Cells(1, 1).Value) ' Producto
            Hoja3.Range("B7").Value = UCase(fila.Cells(1, 2).Value) ' Categoría
            categoria = CStr(fila.Cells(1, 2).Value)
        Next fila
    End If
    
    Set miDiccionario = Nothing
    ' Eliminar el filtro (opcional)
    tabla.Range.AutoFilter

    AsignarValoresSegmentacion nombreProducto, Hoja3.Range("B7").Value
    
    Exit Sub
    
ManejarError:
    ManejadorError ("EditarProducto")
    If Err.Number = 457 Then
        EliminarDuplicadosDeFactTable nombreProducto, "MATERIAL"
        EliminarDuplicadosDeFactTable nombreProducto, "SERVICIO"
        EliminarDuplicadosDeFactTable nombreProducto, "MANO DE OBRA"
        EliminarDuplicadosDeFactTable nombreProducto, "MARGEN"
        LimpiarFormularios
        MsgBox "Los datos del producto fueron reparados, por favor seleccione el producto nuevamente de la lista", vbInformation, "Cálculo incompleto"
    End If
    

    On Error GoTo 0
    
End Sub
'tipoDeCosto=["SERVICIOS","MATERIALES,"MANO_DE_OBRA"]
'elemento= es el nombre del material, el servicio o la mano de obra
Public Sub GuardarDatosMaestros(tipoDeCosto As String, nombreElemento As String, tipoDeMedida As String, _
                                medidaDeCosto As String, precio As Variant, medidaDeUso As String, _
                                cantidadDeUso As Variant, cantidadUnidadesCompradas As Variant)
    Dim tabla As ListObject
    Dim tablaMaestra As String
    Dim nombreColumnaElemento As String
    Dim columnaElemento As Range
    
    
    parametros(1) = tipoDeCosto
    parametros(2) = nombreElemento
    parametros(3) = tipoDeMedida
    parametros(4) = medidaDeCosto
    parametros(5) = precio
    parametros(6) = medidaDeUso
    parametros(7) = cantidadDeUso
    parametros(8) = cantidadUnidadesCompradas
    
    If AMBIENTE = "DESARROLLO" Then
        On Error GoTo 0
    Else
        On Error GoTo ManejarError
    End If
    
    
    If tipoDeCosto = "SERVICIOS" Then
        tablaMaestra = "MaestroServicios"
        nombreColumnaElemento = "SERVICIO"
    ElseIf tipoDeCosto = "MATERIALES" Then
        tablaMaestra = "MaestroMateriales"
        nombreColumnaElemento = "MATERIAL"
    ElseIf tipoDeCosto = "MANO_DE_OBRA" Then
        tablaMaestra = "MaestroManoDeObra"
        nombreColumnaElemento = "MANO DE OBRA"
    End If

    Set tabla = ThisWorkbook.Sheets(tipoDeCosto).ListObjects(tablaMaestra)
    
    'Obtener la fila donde está el elemento
    
    ' Encuentra la columna por nombre
    Set columnaElemento = tabla.HeaderRowRange.Find(nombreColumnaElemento, LookIn:=xlValues, LookAt:=xlWhole)

    ' Verifica si se encontraron las columnas
    If columnaElemento Is Nothing Then
        MsgBox "No se encontró la columna " & nombreColumnaElemento & ".", vbExclamation
        Exit Sub
    End If
    
    ' Aplica el filtro
    tabla.Range.AutoFilter Field:=columnaElemento.Column, Criteria1:=UCase(nombreElemento)
     
    ' Recorrer la fila visible después del filtro y graba datos maestros
    If Not tabla.DataBodyRange.SpecialCells(xlCellTypeVisible) Is Nothing Then
        For Each fila In tabla.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows
            fila.Cells(1, 2).Value = tipoDeMedida
            fila.Cells(1, 3).Value = medidaDeCosto
            fila.Cells(1, 4).Value = cantidadUnidadesCompradas
            fila.Cells(1, 5).Value = medidaDeUso
            fila.Cells(1, 6).Value = cantidadDeUso
            fila.Cells(1, 7).Value = precio
        Next fila
    End If
      
    
    Exit Sub
    
ManejarError:
    ManejadorError "GuardarDatosMaestros", parametros


End Sub
Public Function costoUnitario(tipoDeMedida As String, medidaDeCosto As String, precio As Variant, medidaDeUso As String, cantidadDeUso As Variant, cantidadUnidadesCompradas As Variant) As Variant
    
    Dim costoPorUnidadDeCompra As Variant
    
    parametros(1) = tipoDeMedida
    parametros(2) = medidaDeCosto
    parametros(3) = precio
    parametros(4) = medidaDeUso
    parametros(5) = cantidadDeUso
    parametros(6) = cantidadUnidadesCompradas
    
    If AMBIENTE = "DESARROLLO" Then
        On Error GoTo 0
    Else
        On Error GoTo ManejarError
    End If


    tipoDeMedida = ExtraerTextoAntesDelEspacio(tipoDeMedida)
    ' Llevar medida de costo (Unidad de Costo) a UNO Ej: 8.5$ por 9 metros, a cuanto el metro, a 0.94)
    
    If tipoDeMedida = "PORCENTAJE" Then
        cantidadUnidadesCompradas = 1
    End If
        
    costoPorUnidadDeCompra = precio / cantidadUnidadesCompradas
    
    ' Buscar factor de conversion entre unidad de uso y unidad de compra
    
    
    factor = ObtenerFactor(tipoDeMedida, medidaDeUso, medidaDeCosto)
    
    ' Retorna el cálculo
    If tipoDeMedida = "PORCENTAJE" Then
        cantidadDeUso = cantidadDeUso / 100
    End If
    
    costoUnitario = Round(cantidadDeUso * costoPorUnidadDeCompra * factor, 3)

    Exit Function
ManejarError:
    ManejadorError "CostoUnitario", parametros

End Function


Sub AsignarValoresSegmentacion(nombreProducto As String, categoria As String)
    Dim slcr As SlicerCache
    
    On Error GoTo ManejarError

        Clear_All_Workbook_Slicer_Filters
    
        ' CATEGORIA
        Set slcr = ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_categoriaProductos")
    
        ' Asigna los valores a la segmentación de datos
        For Each slcrItem In slcr.SlicerItems
            If slcrItem.Name = categoria Then
                slcrItem.Selected = True
            Else
                slcrItem.Selected = False
            End If
        Next slcrItem
        
    
        ' PRODUCTO
        Set slcr = ActiveWorkbook.SlicerCaches("SegmentaciónDeDatos_nombreProducto")
    
        ' Asigna los valores a la segmentación de datos
        For Each slcrItem In slcr.SlicerItems
            If slcrItem.Name = nombreProducto Then
                slcrItem.Selected = True
            Else
                slcrItem.Selected = False
            End If
        Next slcrItem
    
    Exit Sub
    
ManejarError:
    ManejadorError ("AsignarValoresSegmentacion")
    On Error GoTo 0

End Sub
Public Sub Clear_All_Workbook_Slicer_Filters()
    
    On Error GoTo ManejarError

    ' Clears all slicer filters in all worksheets
    ' Developed by ExcelPowerTips - visit my Youtube Channel - Like - Subscribe - Share
    
    ' Switch off screen updating for performance
    Application.ScreenUpdating = False
    ' Declare single Slicer Cache type variable
    Dim cache As SlicerCache
    ' Cycle through each cache in whole Excel workbook and clear any manual filter applied
    For Each cache In ActiveWorkbook.SlicerCaches
        cache.ClearManualFilter
    Next cache
    ActiveSheet.Range("B13").Select
    ' Switch back on screen updating
    Application.ScreenUpdating = True
    
    Exit Sub
    
ManejarError:
    ManejadorError ("Clear_All_Workbook_Slicer_Filters")
    On Error GoTo 0
    
End Sub
Function OrdenarDiccionarioPorClave(diccionarioOriginal As Object) As Object
    Dim claves() As Variant
    
    On Error GoTo ManejarError
    
        claves = diccionarioOriginal.Keys
    
        Call OrdenarArray(claves)
    
        Dim diccionarioOrdenado As Object
        Set diccionarioOrdenado = CreateObject("Scripting.Dictionary")
    
        Dim clave As Variant
        For Each clave In claves
            diccionarioOrdenado.Add clave, diccionarioOriginal.Item(clave)
        Next clave
    
        Set OrdenarDiccionarioPorClave = diccionarioOrdenado
    
    Exit Function
    
ManejarError:
    ManejadorError ("OrdenarDiccionarioPorClave")
    On Error GoTo 0
    
End Function

Sub OrdenarArray(arr() As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    
    On Error GoTo ManejarError

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    Exit Sub
    
ManejarError:
    ManejadorError ("OrdenarArray")
    On Error GoTo 0
    
End Sub
Sub AgregarListaDesplegableDesdeTabla()
' Analizar caso cuando la tabla está vacía o agregar una fila DUMMY al inicio
    Dim tabla As ListObject
    Dim columna As ListColumn
    Dim rangoDatos As Range
    Dim celda As Range
    Dim nombresUnicos As Object
    Dim diccionarioOrdenado As Object
    Dim nombre As Variant
    Dim lista As String


    On Error GoTo ManejarError


        ' Establecer referencias a la tabla y la columna
        Set tabla = ThisWorkbook.Sheets("FactTable").ListObjects("FactTable") ' Reemplaza con tus nombres
        Set columna = tabla.ListColumns("nombreProducto") ' Reemplaza con el nombre de tu columna
        Set rangoDatos = columna.DataBodyRange
    
        ' Crear un objeto Dictionary para almacenar nombres únicos
        Set nombresUnicos = CreateObject("Scripting.Dictionary")
        ' Crear un objeto Dictionary para almacenar los nombres ordenados alfabeticamente
        Set diccionarioOrdenado = CreateObject("Scripting.Dictionary")
    
        ' Recorrer el rango de datos y agregar nombres únicos al Dictionary
        For Each celda In rangoDatos
            If Not nombresUnicos.Exists(UCase(celda.Value)) Then
                nombresUnicos.Add UCase(celda.Value), 1
            End If
        Next celda
        
        ' Ordenar el diccionario
        Set diccionarioOrdenado = OrdenarDiccionarioPorClave(nombresUnicos)
    
        ' Crear la cadena de la lista
        For Each nombre In diccionarioOrdenado.Keys
            lista = lista & UCase(nombre) & ","
        Next nombre
    
        ' Eliminar la coma final
        lista = Left(lista, Len(lista) - 1)
    
        ' Establecer la celda donde se aplicará la lista desplegable
        Set celda = ThisWorkbook.Sheets("Formulario").Range("B13")
    
        ' Aplicar la validación de datos de lista desplegable
        With celda.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=lista
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        
    Exit Sub
    
ManejarError:
    ManejadorError ("AgregarListaDesplegableDesdeTabla")
    On Error GoTo 0

End Sub
Sub LlenarComboBoxUnidadCostoUnidadCantidad(tipoDeMedidaConEjemplo As String, tipoDeCosto As String)

    Dim miDiccionario As Object
    Dim valoresOrdenados As Variant
    Dim i As Integer

    ' Llamar a la función y obtener el diccionario
    Set miDiccionario = ObtenerValoresFiltrados("MaestroUnidades", "UNIDADES", "TIPO DE MEDIDA CON EJEMPLO", tipoDeMedidaConEjemplo, "UNIDADES", "ORDENAR")

    ' Ordenar el diccionario en un arreglo bidimensional
    valoresOrdenados = diccionarioOrdenado(miDiccionario)
    
    If UCase(tipoDeCosto) = "SERVICIO" Then
        ' Limpiar el ComboBox antes de agregar nuevos valores
        ServiciosEditar.ComboBoxUnidadCosto.Clear ' Asegúrate de que "servicios" sea el nombre de tu formulario
    
        ' Verificar si se encontraron valores
        If UBound(valoresOrdenados, 1) > 0 Then
            ' Recorrer el arreglo y agregar los valores al ComboBox
            For i = 1 To UBound(valoresOrdenados)
                ServiciosEditar.ComboBoxUnidadCosto.AddItem valoresOrdenados(i, 2) ' Agrega la segunda columna (valor)
                ServiciosEditar.ComboBoxUnidadCantidad.AddItem valoresOrdenados(i, 2) ' Agrega la segunda columna (valor)
            Next i
        Else
            MsgBox "No se encontraron valores que coincidan con los criterios de filtro."
        End If
    ElseIf UCase(tipoDeCosto) = "MATERIAL" Then
        ' Limpiar el ComboBox antes de agregar nuevos valores
        MaterialesEditar.ComboBoxUnidadCosto.Clear ' Asegúrate de que "servicios" sea el nombre de tu formulario
    
        ' Verificar si se encontraron valores
        If UBound(valoresOrdenados, 1) > 0 Then
            ' Recorrer el arreglo y agregar los valores al ComboBox
            For i = 1 To UBound(valoresOrdenados)
                MaterialesEditar.ComboBoxUnidadCosto.AddItem valoresOrdenados(i, 2) ' Agrega la segunda columna (valor)
                MaterialesEditar.ComboBoxUnidadCantidad.AddItem valoresOrdenados(i, 2) ' Agrega la segunda columna (valor)
            Next i
        Else
            MsgBox "No se encontraron valores que coincidan con los criterios de filtro."
        End If
        
    ElseIf UCase(tipoDeCosto) = "MANO DE OBRA" Then
        ' Limpiar el ComboBox antes de agregar nuevos valores
        ManoDeObraEditar.ComboBoxUnidadCosto.Clear ' Asegúrate de que "servicios" sea el nombre de tu formulario
    
        ' Verificar si se encontraron valores
        If UBound(valoresOrdenados, 1) > 0 Then
            ' Recorrer el arreglo y agregar los valores al ComboBox
            For i = 1 To UBound(valoresOrdenados)
                ManoDeObraEditar.ComboBoxUnidadCosto.AddItem valoresOrdenados(i, 2) ' Agrega la segunda columna (valor)
                ManoDeObraEditar.ComboBoxUnidadCantidad.AddItem valoresOrdenados(i, 2) ' Agrega la segunda columna (valor)
            Next i
        Else
            MsgBox "No se encontraron valores que coincidan con los criterios de filtro."
        End If
        
    End If

End Sub
Sub ActualizarTablasDinamicas()
    Dim ws As Worksheet
    Dim pt As PivotTable

    ' Recorre todas las hojas del libro
    For Each ws In ThisWorkbook.Worksheets
        ' Recorre todas las tablas dinámicas en cada hoja
        For Each pt In ws.PivotTables
            ' Actualiza la tabla dinámica
            pt.RefreshTable
        Next pt
    Next ws

End Sub
Sub EscribirCosto(nombreTabla As String, nombreBuscado As String, costoAEscribir As Variant)

    Dim tabla As ListObject
    Dim fila As ListRow

    On Error GoTo ManejarError

        Set tabla = ThisWorkbook.Sheets("Formulario").ListObjects(nombreTabla)
    
    
        ' Recorrer las filas de la tabla
        For Each fila In tabla.ListRows
            ' Verificar si el valor de la primera columna coincide con el nombreBuscado
            If fila.Range.Cells(1, 1).Value = nombreBuscado Then
                ' Escribir el número en la segunda columna
                fila.Range.Cells(1, 2).Value2 = costoAEscribir
                Exit For
            End If
        Next fila
    
    Exit Sub
        
ManejarError:
    ManejadorError ("EscribirCosto")
    On Error GoTo 0

End Sub
Sub ManejadorError(subrutina As String, Optional parametros As Variant = Nothing)

    Dim tabla As ListObject
    Dim cantidadElementos As Integer
    Dim mensajeError As String
    Dim i As Integer
    Dim strParametros As String
          
    
    strParametros = ""
    mensajeError = "Ocurrió un error en: " & Err.Source & vbCrLf & _
           "Descripción: " & Err.Description & vbCrLf & _
           "Número de error: " & Err.Number & ", Subrutina: " & subrutina
           

    If Not parametros Is Nothing Then
         cantidadElementos = UBound(parametros) - LBound(parametros) + 1
    Else
        cantidadElementos = 0
    End If
    
    'Grabar error en la hoja Notas
    If cantidadElementos > 0 Then
        For i = LBound(parametros) To UBound(parametros)
            strParametros = strParametros & "-" & parametros(i)
        Next i
        Debug.Print "MENSAJE DE ERROR: '" & mensajeError & "'; LISTA DE PARAMETROS: '" & strParametros & "' " & Format(Now())
    Else
        Debug.Print "MENSAJE DE ERROR: '" & mensajeError & "' " & Format(Now())
    End If
    
    
    
    Set tabla = ThisWorkbook.Sheets("Notas").ListObjects("Errores")
    
    Set nuevaFila = tabla.ListRows.Add
    nuevaFila.Range.Cells(1, 1).Value = Err.Number
    nuevaFila.Range.Cells(1, 2).Value = "MENSAJE DE ERROR: '" & Err.Description & "'; LISTA DE PARAMETROS: '" & strParametros & "' " & Format(Now())
    nuevaFila.Range.Cells(1, 3).Value = subrutina
    
    Err.Clear ' Borrar el error para evitar comportamientos inesperados
    InicializarVariablesGlobales
    InicializarFactoresDeConversion
    MsgBox "Ocurrió un error, vuelva a intentar", vbInformation, "Cálculo incompleto"
End Sub

Sub ListarTablas()

    Dim hoja As Worksheet
    Dim tabla As ListObject
    Dim mensaje As String

    mensaje = "Lista de tablas en el libro:" & vbCrLf & vbCrLf

    ' Recorrer todas las hojas del libro
    For Each hoja In ThisWorkbook.Worksheets
        ' Recorrer todas las tablas de la hoja
        For Each tabla In hoja.ListObjects
            mensaje = mensaje & "Tabla: " & tabla.Name & " - Hoja: " & hoja.Name & vbCrLf
        Next tabla
    Next hoja

    ' Mostrar el mensaje con la lista de tablas
    Debug.Print mensaje

End Sub

Function CeldaEnTabla(celda As Range, nombreTabla As String, nombreHoja As String) As Boolean

    Dim tabla As ListObject
    Dim filaTabla As ListRow

    On Error Resume Next ' Manejar errores si la celda no está en una tabla

    Set tabla = ThisWorkbook.Sheets(nombreHoja).ListObjects(nombreTabla)
    ' Recorrer todas las filas de la tabla
    For Each filaTabla In tabla.ListRows
        ' Verificar si la celda está dentro del rango de la fila
        If Not Intersect(celda, filaTabla.Range) Is Nothing Then
            ' La celda está en la tabla
            CeldaEnTabla = True
            Exit Function ' Salir de la función
        End If
    Next filaTabla

    CeldaEnTabla = False

End Function



Sub OcultarHoja(nombreHoja As String)

    Dim hoja As Worksheet

    ' Establecer la hoja a ocultar
    Set hoja = ThisWorkbook.Sheets(nombreHoja)

    ' Ocultar la hoja de forma que no se pueda mostrar con clic derecho
    hoja.Visible = xlSheetVeryHidden

    ' Limpiar la variable de objeto
    Set hoja = Nothing

End Sub
Sub MostrarHoja(nombreHoja As String)
    
    Dim hoja As Worksheet
    
    Set hoja = ThisWorkbook.Sheets(nombreHoja)
    hoja.Visible = xlSheetVisible ' O hoja.Visible = True
    Set hoja = Nothing
    
End Sub

Sub ListarSegmentaciones()
    Dim oSlicer As Slicer
    Dim oSlicercache As SlicerCache
    Dim oPT As PivotTable
    Dim oSh As Worksheet


    For Each oSlicercache In ThisWorkbook.SlicerCaches
        For Each oPT In oSlicercache.PivotTables
            oPT.Parent.Activate
            Debug.Print oSlicercache.Name & "," & oPT.Parent.Name
        Next
    Next
End Sub
Function ValoresUnicosColumnaTabla(tabla As ListObject, nombreColumna As String) As Variant

    Dim columna As ListColumn
    Dim rangoDatos As Range
    Dim celda As Range
    Dim dicUnicos As Object
    Dim arregloUnicos() As Variant
    Dim i As Long

    ' Crear un diccionario para almacenar valores únicos
    Set dicUnicos = CreateObject("Scripting.Dictionary")

    ' Encontrar la columna por nombre
    On Error Resume Next
    Set columna = tabla.ListColumns(nombreColumna)
    On Error GoTo 0

    ' Verificar si se encontró la columna
    If columna Is Nothing Then
        ValoresUnicosColumnaTabla = Array() ' Devolver un arreglo vacío si no se encuentra la columna
        Exit Function
    End If

    ' Obtener el rango de datos de la columna
    Set rangoDatos = columna.DataBodyRange

    ' Recorrer las celdas de la columna y agregar valores únicos al diccionario
    For Each celda In rangoDatos
        If Not dicUnicos.Exists(celda.Value) Then
            dicUnicos.Add celda.Value, 1
        End If
    Next celda

    ' Crear el arreglo con los valores únicos
    If dicUnicos.Count > 0 Then
        ReDim arregloUnicos(1 To dicUnicos.Count)
        i = 1
        Dim clave As Variant
        For Each clave In dicUnicos.Keys
            arregloUnicos(i) = clave
            i = i + 1
        Next clave
        ValoresUnicosColumnaTabla = arregloUnicos
    Else
        ValoresUnicosColumnaTabla = Array() ' Devolver un arreglo vacío si no hay valores únicos
    End If

End Function
' Dada una tabla, aplico un filtro sobre una columna y retorna los valores asociados en otra columna
' Hay un parametro opcional para clave del diccionario que permite ordenar, debe indicarse el nombre de la columna de la tabla donde está el orden (con numeros)
Function ObtenerValoresFiltrados(nombreTabla As String, nombreHoja As String, columnaFiltrar As String, valorFiltrar As Variant, columnaBuscar As String, Optional columnaOrdenar As String = "") As Object

    Dim tabla As ListObject
    Dim hoja As Worksheet
    Dim columnaFiltrarObj As ListColumn
    Dim columnaBuscarObj As ListColumn
    Dim columnaOrdenarObj As ListColumn
    Dim fila As ListRow
    Dim dicResultados As Object
    Dim valorBuscar As Variant
    Dim valorOrdenar As Variant

    ' Inicializar el diccionario para almacenar los resultados
    Set dicResultados = CreateObject("Scripting.Dictionary")

    ' Obtener la hoja y la tabla
    Set hoja = ThisWorkbook.Sheets(nombreHoja)
    Set tabla = hoja.ListObjects(nombreTabla)

    ' Obtener las columnas por nombre
    On Error Resume Next
    Set columnaFiltrarObj = tabla.ListColumns(columnaFiltrar)
    Set columnaBuscarObj = tabla.ListColumns(columnaBuscar)
    If columnaOrdenar <> "" Then
        Set columnaOrdenarObj = tabla.ListColumns(columnaOrdenar)
    End If
    On Error GoTo 0

    ' Verificar si se encontraron las columnas
    If columnaFiltrarObj Is Nothing Or columnaBuscarObj Is Nothing Then
        Set ObtenerValoresFiltrados = dicResultados ' Devolver diccionario vacío si no se encuentran las columnas
        Exit Function
    End If

    ' Recorrer las filas de la tabla
    For Each fila In tabla.ListRows
        ' Verificar si el valor de la columna de filtro coincide con el valor a filtrar
        If fila.Range.Cells(1, columnaFiltrarObj.Index).Value = valorFiltrar Then
            ' Obtener el valor de la columna de búsqueda
            valorBuscar = fila.Range.Cells(1, columnaBuscarObj.Index).Value
            ' Si hay columna para ordenar obtener ese valor
            If columnaOrdenar <> "" Then
              valorOrdenar = fila.Range.Cells(1, columnaOrdenarObj.Index).Value
            End If
            ' Agregar el valor al diccionario (la clave y el valor son iguales en caso que no haya columna para ordenar)
            If Not dicResultados.Exists(valorBuscar) Then
                If columnaOrdenar <> "" Then
                    dicResultados.Add valorOrdenar, valorBuscar
                Else
                    dicResultados.Add valorBuscar, valorBuscar
                End If
            End If
        End If
    Next fila

    ' Devolver el diccionario con los valores encontrados
    Set ObtenerValoresFiltrados = dicResultados

End Function
'Ejecuta las acciones para guardar el costo asociado a un material, un servicio o una mano de obra
'Recibe el nombre del material, servicio o personal de mano de obra (descripcionMaterial, descripcionServicio o descripcionManoObra)
'y la unidad base (unidadMaterial, unidadServicio o unidadManoObra)
'la cantidad base se asume 1 (no hay columna en la factTable para esta campo)
'se recibe la cantidad utilizada (cantidadUsadaMaterial, cantidadUsadaServicio, cantidadUsadaManoObra)
'y la unidad de uso (unidadUsadaMaterial, unidadUsadaServicio, unidadUsadaManoObra)
Sub GuardarCostos(nombre As String, unidadBase As String, cantidadUsada As Variant, unidadDeUso As String)



End Sub

' Recibe un diccionario y devuelve un arreglo bidimensional (clave-valor) ordenado
Function diccionarioOrdenado(diccionario As Object) As Variant

    Dim claves() As Variant
    Dim valores() As Variant
    Dim resultado() As Variant
    Dim i As Long, j As Long
    Dim tempClave As Variant
    Dim tempValor As Variant

    ' Obtener las claves y valores del diccionario
    claves = diccionario.Keys
    valores = diccionario.Items

    ' Ordenar las claves alfabéticamente (usando Bubble Sort)
    For i = LBound(claves) To UBound(claves) - 1
        For j = i + 1 To UBound(claves)
            If claves(i) > claves(j) Then
                tempClave = claves(i)
                tempValor = valores(i)
                claves(i) = claves(j)
                valores(i) = valores(j)
                claves(j) = tempClave
                valores(j) = tempValor
            End If
        Next j
    Next i

    ' Crear el arreglo bidimensional con las claves y valores ordenados
    ReDim resultado(1 To diccionario.Count, 1 To 2)
    For i = 1 To diccionario.Count
        resultado(i, 1) = claves(i - 1)
        resultado(i, 2) = valores(i - 1)
    Next i

    ' Devolver el arreglo bidimensional ordenado
    diccionarioOrdenado = resultado

End Function


'Para quitar los espacios en blanco de los tipos de medida con ejemplo
'   textoOriginal = "TIEMPO (EJ:MES)"
'   textoExtraido = "TIEMPO"
Function ExtraerTextoAntesDelEspacio(texto As String) As String
    Dim partes() As String
    Dim resultado As String

    ' Dividir la cadena en partes usando el espacio como delimitador
    partes = Split(texto, " ")

    ' Obtener la primera parte (el texto antes del espacio)
    If UBound(partes) >= 0 Then
        resultado = partes(0)
    Else
        resultado = "" ' Si no hay espacios, devuelve una cadena vacía
    End If

    ExtraerTextoAntesDelEspacio = resultado
End Function



