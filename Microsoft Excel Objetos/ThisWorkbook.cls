VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    ThisWorkbook - Acciones que se realizan al abrir el libro
'------------------------------------------------------------------
'------------------------------------------------------------------
Private Sub Workbook_Open()
    Dim ws As Worksheet
    Dim inicioSheet As Worksheet
    Dim allowedSheets As Variant
    Dim sheetName As String
    Dim isAllowed As Boolean
    Dim i As Integer
    Dim name As Variant
    
    
    ' Lista de nombres permitidos
    allowedSheets = Array("Usuarios", "TECNICO", "Admins", "referenciasHoy", "SUGERENCIAS", "Resultados", "Informantes", "INICIO", "ADMINISTRADOR", "INFORMES")
    
    ' Desactivar la actualización de pantalla para evitar parpadeos
    Application.ScreenUpdating = False
    
    ' Establece la hoja "INICIO"
    Set inicioSheet = ThisWorkbook.Sheets("INICIO")
    
    ' Desoculta la hoja "INICIO"
    inicioSheet.Visible = xlSheetVisible
    
    ' Activa la hoja "INICIO"
    inicioSheet.Activate
    
    ' Oculta todas las demás hojas y elimina las no permitidas
    For i = ThisWorkbook.Sheets.count To 1 Step -1
        Set ws = ThisWorkbook.Sheets(i)
        sheetName = ws.name
        isAllowed = False
        
        ' Verificar si la hoja está en la lista de nombres permitidos
        For Each name In allowedSheets
            If sheetName = name Then
                isAllowed = True
                Exit For
            End If
        Next name
        
        ' Si la hoja no está permitida, eliminarla
        If Not isAllowed Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        ElseIf sheetName <> "INICIO" Then
            ' Si la hoja está permitida pero no es "INICIO", ocultarla
            ws.Visible = xlSheetHidden
        End If
    Next i
    
    AsignarRuta
    
    nombreLibro.GuardarNombreLibro
    
    
    
    
    ' Restaurar la actualización de pantalla
    Application.ScreenUpdating = True
End Sub
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    ThisWorkbook - Acciones que se realizan al cerrar el libro
'------------------------------------------------------------------
'------------------------------------------------------------------


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Desactivar alertas de guardado
    Application.DisplayAlerts = False
    
    ' Marcar el libro como guardado para evitar el mensaje de confirmación
    ThisWorkbook.Saved = True
    
    ' Reactivar alertas de guardado
    Application.DisplayAlerts = True
End Sub

