VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA Resultados
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub CommandButton1_Click() '--BOTON DE RETORNO
    
    ' referenciado a las hojas
    Dim homeSheetName As String
    Dim resultadosSheet As Worksheet
    
    homeSheetName = "INICIO"
    
    Set resultadosSheet = ThisWorkbook.Sheets("Resultados")
    
    ' Verifica si la hoja Home existe
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(homeSheetName)
    On Error GoTo 0
    
    If Not wsHome Is Nothing Then
        
        
        ' Desoculta la hoja si está oculta
        wsHome.Visible = xlSheetVisible
        
        ' Activa la hoja
        wsHome.Activate

        ' Ocultar la hoja actual
        resultadosSheet.Visible = xlSheetHidden
    Else
        MsgBox "La hoja " & homeSheetName & " no existe.", vbExclamation
    End If
End Sub

