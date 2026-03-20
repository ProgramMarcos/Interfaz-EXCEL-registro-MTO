VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA ADMINS
'---    Aquí se registran los Admins
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub CommandButton1_Click() 'RETORNO
    
    'REFERENCIA A LA HOJA
    Dim homeSheetName As String
    homeSheetName = "referenciasHoy"
    
    ' Verifica si la hoja referenciasHoy existe
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(homeSheetName)
    On Error GoTo 0
    
    If Not wsHome Is Nothing Then
        ' Activa la hoja Home
        wsHome.Activate
        ' Guarda el libro de trabajo
        ThisWorkbook.Save
        ' Oculta la hoja Resultados
        Me.Visible = xlSheetVeryHidden
    Else
        MsgBox "La hoja " & homeSheetName & " no existe.", vbExclamation
    End If
End Sub

