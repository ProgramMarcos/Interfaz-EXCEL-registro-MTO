VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA SUGERENCIAS
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub VolverBoton_Click() '--BOTON DE RETORNO
    Dim wsAdministrador As Worksheet
    Dim wsSugerencias As Worksheet
    
    'referencia las hojas
    Set wsAdministrador = ThisWorkbook.Sheets("ADMINISTRADOR")
    Set wsSugerencias = ThisWorkbook.Sheets("SUGERENCIAS")
    
    ' Desocultar la hoja "SUGERENCIAS"
    wsAdministrador.Visible = xlSheetVisible
    
    ' Ocultar la hoja "ADMINISTRADOR"
    wsSugerencias.Visible = xlSheetVeryHidden
End Sub
