VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()

    
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    
    If AMBIENTE = "DESARROLLO" Then
        On Error GoTo 0
    Else
        On Error GoTo ManejarError
    End If
        ' Ejecutar la acción solo si se modifica la celda B13
        If Not Intersect(Target, Range("B13")) Is Nothing Then
            If Hoja3.Range("B13") <> "" Then
                EditarProducto (Range("B13").Value)
            End If
        End If
        
        ' Ejecutar la acción si cambia el valor de una celda de la tabla materiales
        If Not Intersect(Target, Range(columnaNombreMaterial & "4", columnaNombreMaterial & "15")) Is Nothing Then
            If Target.Count = 1 Then
                If Target.Value <> "" Then
                    Hoja3.Hyperlinks.Add Anchor:=Hoja3.Range(columnaNombreMaterial & Target.Row), Address:="", TextToDisplay:=Target.Value
                End If
            End If
        End If
    
        ' Ejecutar la acción si cambia el valor de una celda de la tabla servicios
        If Not Intersect(Target, Range(columnaNombreServicio & "4", columnaNombreServicio & "15")) Is Nothing Then
            If Target.Count = 1 Then
                If Target.Value <> "" Then
                    Hoja3.Hyperlinks.Add Anchor:=Hoja3.Range(columnaNombreServicio & Target.Row), Address:="", TextToDisplay:=Target.Value
                End If
            End If
        End If
    
        ' Ejecutar la acción si cambia el valor de una celda de la tabla mano de obra
        If Not Intersect(Target, Range(columnaNombreMdo & "4", columnaNombreMdo & "15")) Is Nothing Then
            If Target.Count = 1 Then
                If Target.Value <> "" Then
                    Hoja3.Hyperlinks.Add Anchor:=Hoja3.Range(columnaNombreMdo & Target.Row), Address:="", TextToDisplay:=Target.Value
                End If
            End If
        End If
    
    Exit Sub

ManejarError:

    If Err.Number = 1004 Then
        'Target.Clear
        MsgBox "Por favor vuelve a seleccionar de la lista para poder hacer click", vbInformation, "Actualización incompleta"
    End If
    ManejadorError ("Worksheet_Change")
    On Error GoTo 0

    
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    If CeldaEnTabla(Cells(Target.Range.Row, Target.Range.Column), "TablaF_Materiales", "Formulario") Then
        MostrarFormularioMaterialesEditar Target.Name, Target.Range
    ElseIf CeldaEnTabla(Cells(Target.Range.Row, Target.Range.Column), "TablaF_Servicios", "Formulario") Then
        MostrarFormularioServiciosEditar Target.Name, Target.Range
    ElseIf CeldaEnTabla(Cells(Target.Range.Row, Target.Range.Column), "TablaF_ManoDeObra", "Formulario") Then
        MostrarFormularioManoDeObraEditar Target.Name, Target.Range
    End If
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
