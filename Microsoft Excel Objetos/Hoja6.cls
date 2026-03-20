VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA Informantes
'---    Aquí se registran los Informantes
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub CommandButton1_Click() '--BOTON DE RETORNO


    'referencia a la hoja
    Dim homeSheetName As String
    homeSheetName = "referenciasHoy"
    
    ' Verifica si la hoja Home existe
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(homeSheetName)
    On Error GoTo 0
    
    If Not wsHome Is Nothing Then
        ' Activa la hoja de referencias
        wsHome.Activate
        ' Guarda el libro de trabajo
        ThisWorkbook.Save
        ' Oculta la hoja Resultados
        Me.Visible = xlSheetVeryHidden
    Else
        MsgBox "La hoja " & homeSheetName & " no existe.", vbExclamation
    End If
End Sub

