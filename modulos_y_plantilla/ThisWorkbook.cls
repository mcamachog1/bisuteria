VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub Workbook_Open()

    OcultarHoja ("PrecioUnitario")
    OcultarHoja ("FactTable")
    OcultarHoja ("TablasDinamicas")
    'OcultarHoja ("Notas")
    MostrarHoja ("Notas")
    
    AgregarListaDesplegableDesdeTabla
    CrearListaDesplegable "MaestroMateriales", "MATERIALES", "MATERIAL", "Materiales"
    CrearListaDesplegable "MaestroServicios", "SERVICIOS", "SERVICIO", "Servicios"
    CrearListaDesplegable "MaestroManoDeObra", "MANO_DE_OBRA", "MANO DE OBRA", "ManoDeObra"
    CrearListaDesplegable "MaestroCategorias", "CATEGORIAS", "CATEGORIA"
   
    InicializarVariablesGlobales
    InicializarFactoresDeConversion
    
End Sub

