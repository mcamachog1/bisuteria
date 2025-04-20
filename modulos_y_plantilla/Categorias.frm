VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Categorias 
   Caption         =   "Categorias"
   ClientHeight    =   1536
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   5520
   OleObjectBlob   =   "Categorias.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Categorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Cancelar_Click()
   ' Cerrar el formulario
    Unload Me
End Sub

Private Sub Button_OK_Click()
Dim categoria As String

    categoria = Categorias.txtNombre.Value
    InsertarAlFinal "CATEGORIAS", "MaestroCategorias", categoria
    CrearListaDesplegable "MaestroCategorias", "CATEGORIAS", "CATEGORIA"
    
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
