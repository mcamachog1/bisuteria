VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManoDeObraEditar 
   Caption         =   "Servicios"
   ClientHeight    =   4620
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8772.001
   OleObjectBlob   =   "ManoDeObraEditar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ManoDeObraEditar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Button_Cancelar_Click()
   ' Cerrar el formulario
    Unload Me
End Sub

Private Sub Button_OK_Click()
    
    Dim costo As Variant
    
    If AMBIENTE = "DESARROLLO" Then
        On Error GoTo 0
    Else
        On Error GoTo ManejarError
    End If
        
    If txtNombre = "" Or ComboBoxTipoDeMedida = "" Or ComboBoxUnidadCosto = "" Or _
        TextCosto.Value = "" Or ComboBoxUnidadCantidad = "" Or TextCantidad.Value = "" Or _
        txtCantidadPorUnidad = "" Then
       
            MsgBox "Por favor llena todos los campos antes de guardar", vbExclamation, "AResolver"
            Exit Sub
    End If
    
    costo = CostoUnitario(ExtraerTextoAntesDelEspacio(ComboBoxTipoDeMedida), ComboBoxUnidadCosto, TextCosto.Value, ComboBoxUnidadCantidad, TextCantidad.Value, txtCantidadPorUnidad)
    
    EscribirCosto "TablaF_ManoDeObra", txtNombre, costo
    
    GuardarDatosMaestros "MANO_DE_OBRA", txtNombre, ExtraerTextoAntesDelEspacio(ComboBoxTipoDeMedida), ComboBoxUnidadCosto, _
                            TextCosto.Value, ComboBoxUnidadCantidad, TextCantidad.Value, txtCantidadPorUnidad
    
    Unload Me
    
    Exit Sub
    
ManejarError:

    If Err.Number = 13 Then
        MsgBox "Por favor llena todos los campos antes de calcular", vbInformation, "Campos Incompletos"
    End If
    If Err.Number = 424 Then
        MsgBox "Por favor vuelve a calcular", vbInformation, "Cálculo incompleto"
    End If
    ManejadorError ("MaterialesEditar.CommandButtonCalcularCosto_Click")
    On Error GoTo 0

End Sub




Private Sub ButtonEliminar_Click()

    Dim tabla As ListObject
     
    respuesta = MsgBox("¿Está seguro que desea eliminar '" & UCase(txtNombre) & "' de la lista de Mano De Obra?", vbYesNo + vbQuestion, "Confirmar eliminación")
    If respuesta = vbYes Then
        Set tabla = ThisWorkbook.Sheets("MANO_DE_OBRA").ListObjects("MaestroManoDeObra")
        EliminarFilasPorValorColumna txtNombre, "MANO DE OBRA", tabla
        
        Set tabla = Nothing
        Hoja3.Cells(label_row, label_column).Clear
        Hoja3.Cells(label_row, ManoDeObra.Editar.label_column + 1).Clear
        CrearListaDesplegable "MaestroManoDeObra", "MANO_DE_OBRA", "MANO DE OBRA", "MANODEOBRA"
    End If
    
    MsgBox "Registro eliminado", vbInformation, "Éxito"
 
End Sub

Private Sub ComboBoxTipoDeMedida_Change()
    
    LlenarComboBoxUnidadCostoUnidadCantidad ComboBoxTipoDeMedida.Value, "MANO DE OBRA"
    
End Sub

Private Sub ComboBoxUnidadCosto_Change()
    
    If ComboBoxUnidadCosto = "%" Then
        txtCantidadPorUnidad = 100
        txtCantidadPorUnidad.Enabled = False
    Else
        txtCantidadPorUnidad = 1
        txtCantidadPorUnidad.Enabled = True
    End If
End Sub

Public Sub CommandButtonCalcularCosto_Click()
   
    
    If AMBIENTE = "DESARROLLO" Then
        On Error GoTo 0
    Else
        On Error GoTo ManejarError
    End If
    
    If txtNombre = "" Or ComboBoxTipoDeMedida = "" Or ComboBoxUnidadCosto = "" Or _
        TextCosto.Value = "" Or ComboBoxUnidadCantidad = "" Or TextCantidad.Value = "" Or _
        txtCantidadPorUnidad = "" Then
       
            MsgBox "Por favor llena todos los campos antes de calcular", vbExclamation, "AResolver"
            Exit Sub
    End If

    LabelCostoProrrateado.Caption = CostoUnitario(ExtraerTextoAntesDelEspacio(ComboBoxTipoDeMedida), ComboBoxUnidadCosto, TextCosto.Value, ComboBoxUnidadCantidad, TextCantidad.Value, txtCantidadPorUnidad)

    Exit Sub
    
ManejarError:
    ManejadorError ("ManoDeObraEditar.CommandButtonCalcularCosto_Click")
    If Err.Number = 13 Then
        MsgBox "Por favor llena todos los campos antes de calcular", vbInformation, "Campos Incompletos"
    End If
    If Err.Number = 424 Then
        MsgBox "Por favor vuelve a calcular", vbInformation, "Cálculo incompleto"
    End If
    On Error GoTo 0
End Sub

Private Sub LabelCantidad_Click()

End Sub

Private Sub LabelTipoDeMedida_Click()

End Sub

Private Sub LabelUnidad_Click()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim valores As Variant
    Dim i As Long
    Dim cadenaValores As String

    ' Cadena de ejemplo con los valores separados por comas
    cadenaValores = CrearListaDeValores("MaestroUnidades", "UNIDADES", "UNIDADES")

    ' Dividir la cadena en una matriz utilizando la coma como delimitador
    valores = Split(cadenaValores, ",")

    ' Agregar los valores al ComboBox
    For i = LBound(valores) To UBound(valores)
        ComboBoxUnidadCosto.AddItem Trim(valores(i)) ' Usar Trim para eliminar espacios en blanco adicionales
    Next i
End Sub
