VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA USUARIOS
'---    Aquí se registran usuarios
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub CommandButton1_Click() '--BOTON PARA AÑADIR ADMINISTRADORES
    Dim sheetName As String
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim userSheet As Worksheet
    Dim found As Range
    
    'Nombre de la hoja
    sheetName = "Admins"
    
    ' Inicializa la variable global userName
    userName = ""
    
    ' Muestra el UserForm para ingresar el usuario
    UserForm2.Show
    
    ' Verifica si se ingresó un nombre de usuario
    If userName = "" Then
        MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    
    
    ' Verifica si el usuario existe en la columna A de la hoja Admins
    Set userSheet = ThisWorkbook.Sheets("Admins")
    Set found = userSheet.columns("A").Find(What:=userName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If found Is Nothing Then
        MsgBox "Usuario no encontrado. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    
    
    ' Verifica si la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' Obtén la hoja actual
        Set currentSheet = ThisWorkbook.ActiveSheet
        
        ' Oculta la hoja actual
        Me.Visible = xlSheetVeryHidden
        
        ' Desoculta la hoja si está oculta
        ws.Visible = xlSheetVisible
        
        ' Activa la hoja
        ws.Activate
    Else
        'muestra el posible error en pantalla
        MsgBox "La hoja " & sheetName & " no existe.", vbExclamation
    End If
End Sub

Private Sub CommandButton2_Click() '--BOTON DE RETORNO A LA HOJA DE REFERENCIAS
    Dim homeSheetName As String
    homeSheetName = "referenciasHoy"
    
    ' Verifica si la hoja Home existe
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(homeSheetName)
    On Error GoTo 0
    
    If Not wsHome Is Nothing Then
        ' Activa la hoja Home
        wsHome.Activate
        ThisWorkbook.Save
        ' Oculta la hoja Resultados
        Me.Visible = xlSheetVeryHidden
    Else
        'muestra el posible error en pantalla
        MsgBox "La hoja " & homeSheetName & " no existe.", vbExclamation
    End If
End Sub

Private Sub CommandButton3_Click() '--BOTON PARA AÑADIR INFORMANTES
    Dim sheetName As String
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim userSheet As Worksheet
    Dim found As Range
    
    sheetName = "Informantes"
    
    ' Inicializa la variable global userName
    userName = ""
    
    ' Muestra el UserForm para ingresar el usuario
    UserForm2.Show
    
    ' Verifica si se ingresó un nombre de usuario
    If userName = "" Then
        MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    
    
    ' Verifica si el usuario existe en la columna A de la hoja Admins
    Set userSheet = ThisWorkbook.Sheets("Informantes")
    Set found = userSheet.columns("A").Find(What:=userName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If found Is Nothing Then
        MsgBox "Usuario no encontrado. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    
    
    ' Verifica si la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' Obtén la hoja actual
        Set currentSheet = ThisWorkbook.ActiveSheet
        
        ' Oculta la hoja actual
        Me.Visible = xlSheetVeryHidden
        
        ' Desoculta la hoja si está oculta
        ws.Visible = xlSheetVisible
        
        ' Activa la hoja
        ws.Activate
    Else
        'muestra el posible error en pantalla
        MsgBox "La hoja " & sheetName & " no existe.", vbExclamation
    End If
End Sub

