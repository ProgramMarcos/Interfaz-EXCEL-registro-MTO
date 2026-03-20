VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA referenciasHoy
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub CommandButton1_Click()  'importar imagenes
    Dim fDialog As FileDialog
    Dim carpetaSeleccionada As String
    Dim archivo As String
    Dim fila As Long
    Dim nombreArchivoSinExtension As String
    Dim wsRefs As Worksheet
    Dim wsHome As Worksheet
    Dim filePath As String
    Dim robot As String
    Dim tipo As String
    
    'REFERENCIA A LAS HOJAS
    Set wsRefs = ThisWorkbook.Sheets("referenciasHoy")
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    
    ' Crear un cuadro de diálogo para seleccionar la carpeta
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fDialog
        .title = "Seleccionar Carpeta"
        .AllowMultiSelect = False
        If .Show = -1 Then ' Si el usuario selecciona una carpeta
            carpetaSeleccionada = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Encontrar la primera fila vacía en la columna C de la hoja "referenciasHoy", comenzando desde C2
    fila = 2
    Do While wsRefs.Cells(fila, 3).Value <> ""
        fila = fila + 1
    Loop
    
    ' Recorrer todos los archivos de la carpeta seleccionada
    archivo = Dir(carpetaSeleccionada & "\*.*")
    Do While archivo <> ""
        ' Verificar si el archivo es una imagen
        If LCase(Right(archivo, 4)) = ".jpg" Or LCase(Right(archivo, 4)) = ".png" Then
            ' Guardar la dirección del archivo en la columna C
            wsRefs.Cells(fila, 3).Value = carpetaSeleccionada & "\" & archivo
            
            ' Quitar la extensión del nombre del archivo y guardarlo en la columna B
            nombreArchivoSinExtension = Left(archivo, InStrRev(archivo, ".") - 1)
            wsRefs.Cells(fila, 2).Value = nombreArchivoSinExtension
            
            ' Guardar ruta de la carpeta
            wsRefs.Cells(fila, 4).Value = carpetaSeleccionada
            
            'Añadir/Modificar proyecto
            ' Identificar el grupo al que pertenece la pieza
            If Left(nombreArchivoSinExtension, 2) = "P1" Then
                tipo = "P1"
            ElseIf Left(nombreArchivoSinExtension, 4) = "SOLD" Then
                tipo = "SOLD"
            Else
                tipo = "Desconocido"
            End If
            
            ' Guardar el grupo en la columna A
            wsRefs.Cells(fila, 1).Value = tipo
            
            fila = fila + 1
        End If
        archivo = Dir
    Loop
    
    ' Leer el archivo CORDONES.txt y extraer el robot para cada archivo
    filePath = carpetaSeleccionada & "\CORDONES.txt"
    If Dir(filePath) <> "" Then
        Call ExtraerRobots(filePath, wsRefs)
    Else
        MsgBox "El archivo CORDONES.txt no se encuentra en la carpeta", vbExclamation
        Exit Sub
    End If
    
    ' Llenar el ComboBox en la hoja "Home"
    Call RellenarCB
    
    MsgBox "Direcciones de imágenes guardadas en la columna C.", vbInformation
End Sub


Sub RellenarCB()
    Dim lastRow As Long
    Dim i As Long
    Dim wsHome As Worksheet
    Dim wsRefs As Worksheet
    Dim ComboBox As Object
    
    Set wsRefs = ThisWorkbook.Sheets("referenciasHoy")
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    
    ' Encuentra la última fila con datos en la columna B de la hoja "referenciasHoy"
    lastRow = wsRefs.Cells(wsRefs.Rows.count, "B").End(xlUp).Row
    
    ' Referencia al ComboBox ActiveX en la hoja "Home"
    Set ComboBox = wsHome.OLEObjects("ComboBox1").Object
    
    ' Limpia el ComboBox antes de llenarlo
    ComboBox.Clear
    
    ' Llena el ComboBox con los valores de la columna B
    For i = 2 To lastRow ' Asumiendo que los datos comienzan en la fila 2
        ComboBox.AddItem wsRefs.Cells(i, 2).Value
    Next i
End Sub

Sub ExtraerRobots(filePath As String, wsRefs As Worksheet)
    Dim fileContent As String
    Dim lines() As String
    Dim line As Variant
    Dim i As Long
    Dim pieza As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim Data As String
    Dim columns() As String
    Dim parenData As String
    Dim parenArray() As String
    Dim robot As String
    
    ' LEER EL CONTENIDO DEL ARCHIVO
        
    ' Se lee el archivo línea a línea y se guarda en fileContent
    '#1 es el número de archivo o el identificador de archivo.
    'Cuando se abre un archivo con Open filePath For Input As #1,
    'el #1 se utiliza para referirse a ese archivo específico en las
    'operaciones de lectura y escritura que siguen.
    'Es una forma de identificar el archivo que se ha abierto.
    fileContent = ""
    Open filePath For Input As #1
    Do Until EOF(1)
        Line Input #1, line
        fileContent = fileContent & line & vbCrLf
    Loop
    Close #1
    
    ' Buscar cada pieza en el contenido del archivo
    lines = Split(fileContent, vbCrLf)
    For i = 2 To wsRefs.Cells(wsRefs.Rows.count, "B").End(xlUp).Row
        pieza = wsRefs.Cells(i, 2).Value
        robot = "No encontrado"
        
        For Each line In lines
            If InStr(line, pieza) > 0 Then
                startPos = InStr(line, "{") + 1
                endPos = InStr(line, "}")
                Data = Mid(line, startPos, endPos - startPos)
                columns = Split(Data, ";")
                
                ' Extraer la parte entre paréntesis al principio
                If InStr(columns(0), "(") > 0 And InStr(columns(0), ")") > 0 Then
                    parenData = Mid(columns(0), 2, InStr(columns(0), ")") - 2)
                    parenArray = Split(parenData, ",")
                    robot = parenArray(1) ' El robot está en la segunda posición del array
                    Exit For
                End If
            End If
        Next line
        
        wsRefs.Cells(i, 5).Value = robot ' Guardar el robot en la columna E
    Next i
End Sub

Private Sub CommandButton2_Click() 'ocultar y volver a ADMINISTRADOR

    'REFERENCIA A LA HOJA
    Dim homeSheetName As String
    homeSheetName = "ADMINISTRADOR"
    
    ' Verifica si la hoja Home existe
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(homeSheetName)
    On Error GoTo 0
    
    If Not wsHome Is Nothing Then
        ' Desoculta la hoja si está oculta
        wsHome.Visible = xlSheetVisible
        ' Activa la hoja Home
        wsHome.Activate
        ' Oculta la hoja Resultados
        Me.Visible = xlSheetVeryHidden
    Else
        MsgBox "La hoja " & homeSheetName & " no existe.", vbExclamation
    End If
End Sub

Private Sub CommandButton3_Click() 'agregar user
    Dim sheetName As String
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim userSheet As Worksheet
    Dim found As Range
    
    'se referencia la hoja
    sheetName = "Usuarios"
    
    ' Inicializa la variable global userName
    userName = ""
    
    ' Muestra el UserForm para ingresar el usuario
    UserForm2.Show
    
    ' Verifica si se ingresó un nombre de usuario
    If userName = "" Then
        MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    
    
    ' Verifica si el usuario existe en la columna A de la hoja Usuarios
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
        'currentSheet.Visible = xlSheetHidden
        
        ' Desoculta la hoja si está oculta
        ws.Visible = xlSheetVisible
        
        ' Activa la hoja
        ws.Activate
    Else
        MsgBox "La hoja " & sheetName & " no existe.", vbExclamation
    End If
End Sub
