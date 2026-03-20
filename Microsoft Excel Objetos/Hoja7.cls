VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA DE INICIO
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub BotonSugerencias_Click() '--BOTÓN DE SUGERENCIAS
    Sugerencias.Show
End Sub

Private Sub CommandButton1_Click() '--BOTÓN ACCESO A tecnicos

    Dim sheetName As String
    Dim wsInicio As Worksheet
    Dim ws As Worksheet
    Dim userSheet As Worksheet
    Dim foundUser As Range
    Dim currentDate As String
    
    On Error GoTo ErrorHandler
    
    'inicializa las hojas TECNICO e INICIO
    Set wsInicio = ThisWorkbook.Sheets("INICIO")
    sheetName = "TECNICO"
    
    ' Inicializa la variable global userName
    userName = ""
    
    ' Muestra el UserForm para ingresar el usuario
    UserForm2.Show
    
    ' Verifica si se ingresó un nombre de usuario
    If Trim(userName) = "" Then
        MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Verifica si el usuario existe en la columna A de la hoja Usuarios
    Set userSheet = ThisWorkbook.Sheets("Usuarios")
    Set foundUser = userSheet.columns("A").Find(What:=Trim(userName), LookIn:=xlValues, LookAt:=xlWhole)
    
    ' y si no lo encuentra le denega el acceso
    If foundUser Is Nothing Then
        MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Verifica si la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then ' SI EXISTE:
        
        ' Desoculta la hoja si está oculta
        ws.Visible = xlSheetVisible
        
        ' Activa la hoja
        ws.Activate
        
        ' Guarda el usuario en la celda G14
        ws.Range("G14").Value = userName
        
        ' Obtiene la fecha actual en formato día/mes/año
        currentDate = Format(Date, "dd/mm/yyyy")
        
        ' Guarda la fecha actual en la celda A14
        ws.Range("A14").Value = currentDate
        
        
        
        ' Ocultar la hoja actual
        wsInicio.Visible = xlSheetHidden

    Else
        MsgBox "La hoja " & sheetName & " no existe.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Se produjo un error: " & Err.Description, vbCritical 'MUESTRA EL ERROR QUE SE HA PODIDO PRODUCIR
End Sub




Private Sub CommandButton2_Click() '--BOTON INFORMES
    Dim sheetName As String
    Dim wsInicio As Worksheet
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim userSheet As Worksheet
    Dim foundUser As Range
    Dim currentDate As String
    
    ' Inicializa las hojas
    Set wsInicio = ThisWorkbook.Sheets("INICIO")
    sheetName = "INFORMES"
    
    ' Inicializa la variable global userName
    userName = ""
    
    ' Muestra el UserForm para ingresar el usuario
    UserForm2.Show
    
    ' Verifica si se ingresó un nombre de usuario
    If Trim(userName) = "" Then
        MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Verifica si el usuario existe en la columna A de la hoja Usuarios
    Set userSheet = ThisWorkbook.Sheets("Informantes")
    Set foundUser = userSheet.columns("A").Find(What:=Trim(userName), LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundUser Is Nothing Then
        MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Verifica si la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        
        ' Desoculta la hoja si está oculta
        ws.Visible = xlSheetVisible
        
        ' Activa la hoja
        ws.Activate
        'guardo el usuario en la celda G12
        ws.Range("G12").Value = userName
        
        ' Obtiene la fecha actual en formato día/mes/año
        currentDate = Format(Date, "dd/mm/yyyy")
        
        ' Guarda la fecha actual en la celda A14
        ws.Range("A14").Value = currentDate
        
        ' Ocultar la hoja actual
        wsInicio.Visible = xlSheetHidden
    Else
        'muestra el posible error en pantalla
        MsgBox "La hoja " & sheetName & " no existe.", vbExclamation
    End If
End Sub


Private Sub CommandButton3_Click() '--BOTON ASDMINISTRADOR
    Dim sheetName As String
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim userSheet As Worksheet
    Dim foundUser As Range
    Dim wsInicio As Worksheet
    
    'Inicializa las hojas
    Set wsInicio = ThisWorkbook.Sheets("INICIO")
    sheetName = "ADMINISTRADOR"
    
    ' Inicializa la variable global userName
    userName = ""
    
    ' Muestra el UserForm para ingresar el usuario
    UserForm2.Show
    
    ' Verifica si se ingresó un nombre de usuario
    If Trim(userName) = "" Then
        MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Verifica si el usuario existe en la columna A de la hoja Usuarios
    Set userSheet = ThisWorkbook.Sheets("Admins")
    Set foundUser = userSheet.columns("A").Find(What:=Trim(userName), LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundUser Is Nothing Then
        MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Verifica si la hoja existe
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ' Obtén la hoja actual
        Set currentSheet = ThisWorkbook.ActiveSheet
        
        ' Desoculta la hoja si está oculta
        ws.Visible = xlSheetVisible
        
        ' Activa la hoja
        ws.Activate
        
        ' Obtiene la fecha actual en formato día/mes/año
        currentDate = Format(Date, "dd/mm/yyyy")
        
        ' Guarda la fecha actual en la celda A14
        ws.Range("A14").Value = currentDate
        
        ' Ocultar la hoja actual
        wsInicio.Visible = xlSheetHidden
    Else
        'muestra el posible error en pantalla
        MsgBox "La hoja " & sheetName & " no existe.", vbExclamation
    End If
End Sub
