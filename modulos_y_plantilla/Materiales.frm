VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Materiales 
   Caption         =   "Materiales"
   ClientHeight    =   1920
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6900
   OleObjectBlob   =   "Materiales.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Materiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Button_Cancelar_Click()
   ' Cerrar el formulario
    Unload Me
End Sub

Private Sub Button_OK_Click()
    Dim material As String
    ' Validar que el material es nuevo. Si existe enviar mensaje de advertencia
    
    material = Materiales.txtNombre.Value
    ' Insertar el material
    InsertarAlFinal "MATERIALES", "MaestroMateriales", material
    
    CrearListaDesplegable "MaestroMateriales", "MATERIALES", "MATERIAL", "Materiales"
    
    Unload Me


End Sub




