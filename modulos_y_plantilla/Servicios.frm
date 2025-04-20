VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Servicios 
   Caption         =   "Servicios"
   ClientHeight    =   1920
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6900
   OleObjectBlob   =   "Servicios.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Servicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Button_Cancelar_Click()
   ' Cerrar el formulario
    Unload Me
End Sub

Private Sub Button_OK_Click()
    Dim servicio As String
    ' Validar que el material es nuevo. Si existe enviar mensaje de advertencia
    
    servicio = Servicios.txtNombre.Value
    ' Insertar el material
    InsertarAlFinal "SERVICIOS", "MaestroServicios", servicio
    
    CrearListaDesplegable "MaestroServicios", "SERVICIOS", "SERVICIO", "Servicios"
    
    Unload Me


End Sub




