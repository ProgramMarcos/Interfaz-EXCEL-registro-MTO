VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA TECNICO
'---    Aquí se generan los reportes
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub Worksheet_Change(ByVal Target As Range)  'llama a la función que cambia la carpeta destino según la fecha
    ' Verifica si la celda que cambió es A14
    If Not Intersect(Target, Me.Range("A14")) Is Nothing Then
        ' Llama a la macro que deseas ejecutar
        Call CambiaCarpetaDestino
    End If
End Sub

Sub CambiaCarpetaDestino()
    Dim fechaCarpeta As String
    Dim carpetaFecha As String
    Dim fileName As String
    Dim count As Integer
    Dim ws As Worksheet
    
    ' Asigna la hoja "TECNICO" a la variable ws
    Set ws = ThisWorkbook.Sheets("TECNICO")
    
    ' Inicializa el contador
    count = 0
    
    ' Manejo de errores
    On Error Resume Next
    
    ' Extraer y formatear la fecha de la celda A14
    fechaCarpeta = Format(CDate(ws.Range("A14").Value), "yyyy_mm_dd")
    
    ' Verificar si ocurrió un error
    If Err.Number <> 0 Then
        ' Restaurar las configuraciones de Excel
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        ws.Range("G6").Value = count
        ' Limpiar el error
        Err.Clear
        Exit Sub
    End If
    
    ' Restaurar el manejo de errores
    On Error GoTo 0
    
    ' Definir la ruta de la carpeta con la fecha
    'carpetaFecha = ThisWorkbook.Path & "\Results\" & fechaCarpeta & "\"
    ruta.AsignarRuta
    carpetaFecha = ruta.ruta & "\" & fechaCarpeta & "\"
    
    
    ' Verificar si la carpeta con la fecha existe
    If Dir(carpetaFecha, vbDirectory) = "" Then
        ' Restaurar las configuraciones de Excel
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        glblCount.modCount (count)
        Exit Sub
    Else
        fileName = Dir(carpetaFecha & "*.xlsx")
        Do While fileName <> ""
            count = count + 1
            fileName = Dir() ' Obtener el siguiente archivo
        Loop
        'ws.Range("G6").Value = count
        glblCount.modCount (count)
    End If
End Sub



Private Sub Worksheet_Activate() 'Acciones que se realizan al activar la hoja TECNICO
    Dim fechaCarpeta As String
    Dim carpetaFecha As String
    Dim fileName As String
    

    Dim count As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TECNICO")
    
    ' Referencia al ComboBox ActiveX
    Dim cb As Object
    Set cb = ws.OLEObjects("ComboBox3").Object
    
    ' Limpiar cualquier elemento existente en el ComboBox
    cb.Clear
    
    
    '--SE RELLENA EL COMBOBOX DE FILTRAR PROYECTO CON LOS PROYECTOS DESEADOS
    
    'Añadir/Modificar proyecto
    ' Agregar las opciones deseadas al ComboBox
    cb.AddItem " "
    cb.AddItem "P1"
    cb.AddItem "SOLD"
 
    'SE OBTIENE LA RUTA DE LA CARPETA DESTINO
    count = 0
    ' Extraer y formatear la fecha de la celda A14
    fechaCarpeta = Format(CDate(ws.Range("A14").Value), "yyyy_mm_dd")
    
    ' Definir la ruta de la carpeta con la fecha
    ruta.AsignarRuta
    carpetaFecha = ruta.ruta & "\" & fechaCarpeta & "\"
    'carpetaFecha = ThisWorkbook.Path & "\Results\" & fechaCarpeta & "\"
    
    ' Verificar si la carpeta con la fecha existe
    If Dir(carpetaFecha, vbDirectory) = "" Then
        
        ' Restaurar las configuraciones de Excel
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        glblCount.modCount (count)
        Exit Sub
        
    Else
        fileName = Dir(carpetaFecha & "*.xlsx")
        Do While fileName <> ""
            count = count + 1
            fileName = Dir() ' Obtener el siguiente archivo
        Loop
        glblCount.modCount (count)
    End If
    
End Sub

Private Sub ComboBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'RUTINA PARA MOSTRAR LOS REPORTES REALIZADOS EN EL COMBOBOX INFERIOR
    Dim carpetaResultados As String
    Dim archivo As String
    Dim fechaCarpeta As String
    Dim carpetaFecha As String
    Dim wsHome As Worksheet
    Dim count As Integer
    
    count = 0
    
    ' Asigna la hoja "TECNICO"
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    
    ' Extraer y formatear la fecha de la celda A14
    fechaCarpeta = Format(CDate(wsHome.Range("A14").Value), "yyyy_mm_dd")
    
    ' Definir la ruta de la carpeta con la fecha
    'carpetaFecha = ThisWorkbook.Path & "\Results\" & fechaCarpeta
    ruta.AsignarRuta
    carpetaFecha = ruta.ruta & "\" & fechaCarpeta & "\"
    'carpetaFecha = "\\vigmsnas\Industrial-Development\02 - SOLDADURA\0_INDICE\24_COMPARTIDA MARCOS LOPEZ\3. TASK FORCE\0. Nuevo Programa Registro Regulaciones\Registro regs\Results\" & fechaCarpeta & "\"
    
    
    ' Verificar si la carpeta con la fecha existe, si no, crearla
    
    If Dir(carpetaFecha, vbDirectory) = "" Then
        MkDir carpetaFecha
        'MsgBox "La carpeta no existía y ha sido creada.", vbInformation
    End If
    
    ComboBox2.Clear ' Limpiar el ComboBox antes de llenarlo
    
    ' Recorrer los archivos en la carpeta con la fecha
    archivo = Dir(carpetaFecha & "\*.xlsx")
    Do While archivo <> ""
        count = count + 1
        ComboBox2.AddItem archivo
        archivo = Dir
    Loop
    glblCount.modCount (count)
End Sub


Private Sub CommandButton1_Click() '--BOTON CARGAR CROQUIS
    Call RellenarCB 'Llamada a la rutina
End Sub

Sub RellenarCB()
    Dim lastRow As Long
    Dim i As Long
    Dim wsTecnico As Worksheet
    Dim wsRefs As Worksheet
    Dim ComboBox As Object
    
    Set wsRefs = ThisWorkbook.Sheets("referenciasHoy")
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    
    ' Encuentra la última fila con datos en la columna B de la hoja "referenciasHoy"
    lastRow = wsRefs.Cells(wsRefs.Rows.count, "B").End(xlUp).Row
    
    ' Referencia al ComboBox ActiveX en la hoja "TECNICO"
    Set ComboBox = wsHome.OLEObjects("ComboBox1").Object
    
    ' Limpia el ComboBox antes de llenarlo
    ComboBox.Clear
    
    ' Llena el ComboBox con los valores de la columna B
    For i = 2 To lastRow ' Asumiendo que los datos comienzan en la fila 2
        ComboBox.AddItem wsRefs.Cells(i, 2).Value
    Next i
End Sub


Private Sub CommandButton2_Click() '--BOTON OK

    Dim sheetName As String
    Dim wsRefs As Worksheet
    Dim wsHome As Worksheet
    Dim wsSeleccionada As Worksheet
    Dim userSheet As Worksheet
    Dim ComboBox As Object
    Dim pieza As String
    Dim puesto As String
    Dim fecha As String
    Dim hora As String
    Dim tipo As String
    Dim hojaAbrir As String
    Dim direccion As String
    Dim CARPETA As String
    Dim filePath As String
    Dim fileContent As String
    Dim found As Boolean
    Dim lines() As String
    Dim line As Variant
    Dim startPos As Integer
    Dim endPos As Integer
    Dim Data As String
    Dim columns() As String
    Dim parenData As String
    Dim parenArray() As String
    Dim mesas() As String
    Dim title As String
    Dim columnData() As String
    Dim oleObj As OLEObject
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim foundUser As Range
    
    Dim fechaCarpeta As String
    Dim carpetaFecha As String
    Dim fileName As String
    
    Dim count As Integer
    Dim newWorkbook As Workbook
    
    ' Desactivar la actualización de pantalla y el cálculo automático
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo CleanUp ' Asegurarse de restaurar la configuración en caso de error
    
    Set wsRefs = ThisWorkbook.Sheets("referenciasHoy")
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    Set userSheet = ThisWorkbook.Sheets("Usuarios")
    Set informerSheet = ThisWorkbook.Sheets("Informantes")
    Set ComboBox = wsHome.OLEObjects("ComboBox1").Object 'combo box de los croquis
    
    ' Verificar si el usuario en G14 está en la columna A de la hoja Usuarios
    If Trim(wsHome.Range("G14").Value) = "" Then
        UserForm2.Show
        If Trim(userName) = "" Then
            MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
            GoTo CleanUp 'Si no lo encuentra se va a otra zona del código
        Else
            Set foundUser = userSheet.columns("A").Find(What:=Trim(userName), LookIn:=xlValues, LookAt:=xlWhole)
            If foundUser Is Nothing Then
                MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
                GoTo CleanUp 'Si no lo encuentra se va a otra zona del código
            Else
                wsHome.Range("G14").Value = userName
            End If
        End If
        
    Else 'SI NO LO ESTÁ BUSCA A VER SI ES UN USER VÁLIDO, UNA VEZ VERIFICADO LO PONE EN A14
        Set foundUser = userSheet.columns("A").Find(What:=wsHome.Range("G14").Value, LookIn:=xlValues, LookAt:=xlWhole)
        If foundUser Is Nothing Then
            UserForm2.Show
            If Trim(userName) = "" Then
                MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
                GoTo CleanUp 'Si no lo encuentra se va a otra zona del código
            Else
                Set foundUser = userSheet.columns("A").Find(What:=Trim(userName), LookIn:=xlValues, LookAt:=xlWhole)
                If foundUser Is Nothing Then
                    MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
                    GoTo CleanUp 'Si no lo encuentra se va a otra zona del código
                Else
                    wsHome.Range("G14").Value = userName
                End If
            End If
        End If
    End If
    
    ' Compruebo el valor del comboBox
    If ComboBox.Value = "" Then
        MsgBox "Debe introducir una pieza.", vbExclamation
        GoTo CleanUp 'Si no la encuentra se va a otra zona del código
    End If
    
    pieza = ComboBox.Value
    fecha = Format(Now, "dd/mm/yyyy")
    hora = Format(Now, "HH:MM:SS")
    
    'Añadir/Modificar proyecto
    ' Identificar proyecto
    If Left(pieza, 2) = "P1" Then
        tipo = "P1"
    ElseIf Left(pieza, 4) = "SOLD" Then
        tipo = "SOLD"
    Else
        MsgBox "Tipo de pieza no identificado.", vbExclamation
        GoTo CleanUp 'Si no la encuentra se va a otra zona del código
    End If
    
    i = 0
    ' Buscar la pieza en la hoja de referencias
    Do While wsRefs.Range("B2").Offset(i, 0).Value <> pieza
        i = i + 1
        ' Evitar bucle infinito
        If wsRefs.Range("B2").Offset(i, 0).Value = "" Then
            MsgBox "Pieza no encontrada.", vbExclamation
            GoTo CleanUp 'Si no la encuentra se va a otra zona del código
        End If
    Loop
    
    hojaAbrir = LimpiarNombreHoja(pieza & " " & hora)
    direccion = wsRefs.Range("B2").Offset(i, 1).Value
    CARPETA = wsRefs.Range("B2").Offset(i, 2).Value
    
    
    Dim blockMod As String
    
    'crea un archivo que impide guardar
    blockMod = ruta.ruta & "\lock.txt"
        
    If Dir(blockMod) <> "" Then
        MsgBox "El archivo está siendo modificado por otro usuario. Por favor, inténtelo más tarde."
        Exit Sub
    End If
    ' Create lock file
    Open blockMod For Output As #1
    Print #1, "Locked"
    Close #1
    
    ' Crear un nuevo libro de Excel
    Set newWorkbook = Workbooks.Add
    Set wsSeleccionada = newWorkbook.Sheets(1)
    wsSeleccionada.name = hojaAbrir
    
    
    ' Llamar a la subrutina para insertar la imagen en el nuevo libro
    Call CrearHojaConImagen(newWorkbook, hojaAbrir, direccion)
    
    ' Verificar si la hoja existe
    On Error Resume Next
    Set wsSeleccionada = newWorkbook.Sheets(hojaAbrir)
    On Error GoTo 0
    
    If wsSeleccionada Is Nothing Then
        MsgBox "La hoja " & hojaAbrir & " no existe.", vbExclamation
    Else
        count = 0
        ' Extraer y formatear la fecha de la celda A14
        fechaCarpeta = Format(CDate(wsHome.Range("A14").Value), "yyyy_mm_dd")
        
        ' Definir la ruta de la carpeta con la fecha
        'carpetaFecha = ThisWorkbook.Path & "\Results\" & fechaCarpeta & "\"
        ruta.AsignarRuta
        carpetaFecha = ruta.ruta & "\" & fechaCarpeta & "\"
    
        
        ' Verificar si la carpeta con la fecha existe
        If Dir(carpetaFecha, vbDirectory) = "" Then
            count = 0
        Else
            fileName = Dir(carpetaFecha & "*.xlsx")
            Do While fileName <> ""
                count = count + 1
                fileName = Dir() ' Obtener el siguiente archivo
            Loop
            glblCount.modCount (count)
            'wsHome.Range("G6").Value = count
        End If
        For Each ws In ThisWorkbook.Worksheets
            If ws.name <> "referenciasHoy" And ws.name <> "Home" And ws.name <> "Resultados" And ws.name <> "Usuarios" And ws.name <> "Admins" And ws.name <> "Informantes" _
                And ws.name <> "TECNICO" And ws.name <> "INFORMES" And ws.name <> "ADMINISTRADOR" And ws.name <> "INICIO" And ws.name <> "SUGERENCIAS" Then
                ComboBox2.AddItem ws.name
                'count = count + 1 ' Incrementar el contador por cada hoja añadida
            End If
        Next ws

        ' Activar la hoja seleccionada
        'Guardar el numero de regulaciones A3
        With wsSeleccionada.Range("A3")
            .Value = glblCount.glblCount + 1
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 255)
        End With
        ' Activar la hoja seleccionada
        
        'Poner "Hora" en A4
        With wsSeleccionada.Range("A4")
            .Value = "Hora :"
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(211, 211, 211)
        End With
        'Espacio para la hora en B4
        With wsSeleccionada.Range("B4")
            .Value = ""
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(211, 211, 211)
        End With
        'Guardar el título en D1
        With wsSeleccionada.Range("D1")
            .Value = pieza
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 255)
            .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
        End With
        'Guardar texto en D2
        With wsSeleccionada.Range("D2")
            .NumberFormat = "@" ' Establecer el formato de la celda como texto
            .Value = ""
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 255)
            .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
        End With
        'Guardar fecha en B3
        With wsSeleccionada.Range("B3")
            .NumberFormat = "@" ' Establecer el formato de la celda como texto
            .Value = fecha
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 255)
            .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
        End With
        ' Guardar el tipo en la celda B1
        With wsSeleccionada.Range("B1")
            .Value = tipo
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 255)
            .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
        End With
        ' Establecer todos los checkboxes a False
        For Each oleObj In wsSeleccionada.OLEObjects
            If TypeName(oleObj.Object) = "CheckBox" Then
                oleObj.Object.Value = False
            End If
        Next oleObj
        
        ' Leer el archivo CORDONES.txt
        filePath = CARPETA & "\CORDONES.txt"
        If Dir(filePath) = "" Then
            MsgBox "El archivo CORDONES.txt no se encuentra en la carpeta", vbExclamation
            GoTo CleanUp 'Si no lo encuentra se va a otra zona del código
        End If
        
        ' LEER EL CONTENIDO DEL ARCHIVO
        
        ' Se lee el archivo línea a línea y se guarda en fileContent
        '#1 es el número de archivo o el identificador de archivo.
        'Cuando se abre un archivo con Open filePath For Input As #1,
        'el #1 se utiliza para referirse a ese archivo específico en las
        'operaciones de lectura y escritura que siguen.
        'Es una forma de identificar el archivo que se ha abierto.
        
        fileContent = ""
        Open filePath For Input As #1
        Do Until EOF(1) 'Repite hasta fin del archivo (End Of File)
            Line Input #1, line
            fileContent = fileContent & line & vbCrLf 'Añade la línea leída a fileContent, seguida de un salto de línea (vbCrLf)
        Loop
        Close #1
        
        ' Buscar la pieza en el contenido del archivo
        lines = Split(fileContent, vbCrLf)
        found = False
        For Each line In lines
            If InStr(line, pieza) > 0 Then
                found = True
                startPos = InStr(line, "{") + 1
                endPos = InStr(line, "}")
                Data = Mid(line, startPos, endPos - startPos)
                columns = Split(Data, ";")
                
                ' Extraer la parte entre paréntesis al principio
                If InStr(columns(0), "(") > 0 And InStr(columns(0), ")") > 0 Then
                    parenData = Mid(columns(0), 2, InStr(columns(0), ")") - 2)
                    parenArray = Split(parenData, ",")
                    columns(0) = Mid(columns(0), InStr(columns(0), ")") + 2)
                    
                    ' Guardar el robot en la celda C2
                    With wsSeleccionada.Range("C2")
                        .Value = parenArray(1)
                        .Font.Size = 14
                        .Font.Bold = True
                        .Interior.color = RGB(200, 200, 255)
                        .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
                    End With
                    
                    ' Guardar el puesto en la celda C1
                    With wsSeleccionada.Range("C1")
                        .Value = parenArray(0)
                        .Font.Size = 14
                        .Font.Bold = True
                        .Interior.color = RGB(200, 200, 255)
                        .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
                    End With
                    
                    ' Guardar el tipo de soldadura en la celda B2
                    With wsSeleccionada.Range("B2")
                        .Value = parenArray(4)
                        .Font.Size = 14
                        .Font.Bold = True
                        .Interior.color = RGB(200, 200, 255)
                        .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
                    End With
                    
                    ' Guardar el usuario en la celda C3
                    With wsSeleccionada.Range("C3")
                        .Value = wsHome.Range("G14").Value
                        .Font.Size = 14
                        .Font.Bold = True
                        .Interior.color = RGB(200, 200, 255)
                        .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
                    End With
                    '  celda C4
                    With wsSeleccionada.Range("D3")
                        .Value = "AVISO :"
                        .Font.Size = 14
                        .Font.Bold = True
                        .Interior.color = RGB(200, 200, 255)
                        .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
                    End With
                    
                    ' Guardar las mesas en un array
                    mesas = Split(parenArray(2) & "," & parenArray(3), ",")
                    If Trim(mesas(0)) <> "" And Trim(mesas(1)) <> "" Then
                        ReDim Preserve mesas(UBound(mesas) + 1)
                        mesas(2) = "M1/M2"
                    End If
                End If
                
                ' Crear las columnas y rellenar los datos
                For j = LBound(columns) To UBound(columns)
                    If InStr(columns(j), ":") > 0 Then
                        title = Mid(columns(j), InStr(columns(j), """") + 1, InStrRev(columns(j), """") - InStr(columns(j), """") - 1)
                        columnData = Split(Split(columns(j), ":")(1), ",")
                        
                        ' Poner el título en la columna
                        With wsSeleccionada.Cells(1, j + 9)
                            .Value = title 'COLUMNA M
                            .Font.Size = 14
                            .Font.Bold = True
                            .Interior.color = RGB(200, 200, 255)
                            .EntireColumn.AutoFit ' Ajustar el ancho de la columna al contenido
                        End With
                        
                        ' Ajustar el ancho de la columna al contenido
                        wsSeleccionada.columns(j + 9).AutoFit
                        If wsSeleccionada.columns(j + 9).ColumnWidth < 70 Then
                            ' Ajustar el ancho de la columna al contenido
                            wsSeleccionada.columns(j + 9).ColumnWidth = 100 ' Ajustar el ancho de la columna
                        End If
                        
                        ' Ajustar la altura de las filas
                        wsSeleccionada.Rows("2:" & UBound(columnData) + 2).RowHeight = 20
                        
                        For k = LBound(columnData) To UBound(columnData) 'desde el borde izquierdo de la columna hasta el borde derecho
                            If columnData(k) <> "}" Then
                                ' Añadir un CheckBox en la celda con el texto
                                Set oleObj = wsSeleccionada.OLEObjects.Add(ClassType:="Forms.CheckBox.1", Link:=False, DisplayAsIcon:=False, _
                                    Left:=wsSeleccionada.Cells(k + 2, j + 9).Left, Top:=wsSeleccionada.Cells(k + 2, j + 9).Top, _
                                    Width:=wsSeleccionada.Cells(k + 2, j + 9).Width * 13 / 100, Height:=wsSeleccionada.Cells(k + 2, j + 9).Height)
                                oleObj.Object.Caption = columnData(k)
                                
                                ' Añadir el cuadro combinado de las mesas al lado del CheckBox
                                Set oleObj = wsSeleccionada.OLEObjects.Add(ClassType:="Forms.ComboBox.1", Link:=False, DisplayAsIcon:=False, _
                                    Left:=wsSeleccionada.Cells(k + 2, j + 9).Left + wsSeleccionada.Cells(k + 2, j + 9).Width * 13 / 100, Top:=wsSeleccionada.Cells(k + 2, j + 9).Top, _
                                    Width:=wsSeleccionada.Cells(k + 2, j + 9).Width * 14 / 100, Height:=wsSeleccionada.Cells(k + 2, j + 9).Height)
                                
                                With oleObj.Object
                                    .List = mesas
                                    .Text = "Mesa"
                                    .ForeColor = RGB(0, 0, 0)
                                    .ColumnCount = 1
                                    .ColumnWidths = "150 pt" ' Ajusta el ancho de la columna para que el contenido se vea completo
                                End With
                                
                                ' Añadir el primer cuadro combinado al lado del CheckBox
                                Set oleObj = wsSeleccionada.OLEObjects.Add(ClassType:="Forms.ComboBox.1", Link:=False, DisplayAsIcon:=False, _
                                    Left:=wsSeleccionada.Cells(k + 2, j + 9).Left + wsSeleccionada.Cells(k + 2, j + 9).Width * 27 / 100, Top:=wsSeleccionada.Cells(k + 2, j + 9).Top, _
                                    Width:=wsSeleccionada.Cells(k + 2, j + 9).Width * (73 / 3) / 100, Height:=wsSeleccionada.Cells(k + 2, j + 9).Height)
                                
                                With oleObj.Object
                                    .List = Array(" ", "Perforaciones", "Desplazamiento", "Exceso de calado", "GAP", "Mal posicionado", "No repite posición", "Corta", "Petardeos", "Rotura al límite")
                                    .Text = "PROBLEMA"
                                    .ForeColor = RGB(0, 0, 0)
                                    .ColumnCount = 1
                                    .ColumnWidths = "150 pt" ' Ajusta el ancho de la columna para que el contenido se vea completo
                                End With
                                
                                ' Añadir el segundo cuadro combinado al lado del primer cuadro combinado
                                Set oleObj = wsSeleccionada.OLEObjects.Add(ClassType:="Forms.ComboBox.1", Link:=False, DisplayAsIcon:=False, _
                                    Left:=wsSeleccionada.Cells(k + 2, j + 9).Left + wsSeleccionada.Cells(k + 2, j + 9).Width * (154 / 3) / 100, Top:=wsSeleccionada.Cells(k + 2, j + 9).Top, _
                                    Width:=wsSeleccionada.Cells(k + 2, j + 9).Width * (73 / 3) / 100, Height:=wsSeleccionada.Cells(k + 2, j + 9).Height)
                                
                                With oleObj.Object
                                    .List = Array(" ", "Mal posicionado", "Parametros", "TCP NOK", "Cuello NOK", "Boquilla desgastada", "GAP", "Utillaje NOK", "Piezas entrantes NOK", "Alimentación hilo NOK", "Suciedad en utillaje", "Masa no conecta", "Limpieza boquilla NOK", "No soplado en posición limpieza", "Tobera con suciedad")
                                    .Text = "CAUSA"
                                    .ForeColor = RGB(0, 0, 0)
                                    .ColumnCount = 1
                                    .ColumnWidths = "150 pt"
                                End With
                                
                                ' Añadir el tercer cuadro combinado al lado del segundo cuadro combinado
                                Set oleObj = wsSeleccionada.OLEObjects.Add(ClassType:="Forms.ComboBox.1", Link:=False, DisplayAsIcon:=False, _
                                    Left:=wsSeleccionada.Cells(k + 2, j + 9).Left + wsSeleccionada.Cells(k + 2, j + 9).Width * (227 / 3) / 100, Top:=wsSeleccionada.Cells(k + 2, j + 9).Top, _
                                    Width:=wsSeleccionada.Cells(k + 2, j + 9).Width * (73 / 3) / 100, Height:=wsSeleccionada.Cells(k + 2, j + 9).Height)
                                
                                With oleObj.Object
                                    .List = Array(" ", "Mover cordón", "Modificar TCP", "Cambiar boquilla", "Cambiar cuello", "Ajustar utillaje", "Cambiar lote pieza", "Igualar parámetros a estandar", "Cambiar rodillos", "Cambiar sirga", "Cambiar fresa limpieza", "Modificar posición limpieza")
                                    .Text = "SOLUCIÓN"
                                    .ForeColor = RGB(0, 0, 0)
                                    .ColumnCount = 1
                                    .ColumnWidths = "150 pt"
                                End With

                            Else
                                Exit For
                            End If
                        Next k

                    End If
                Next j
                
                ' Insertar botones en A1 y A2
                InsertarBotones wsSeleccionada 'subrutina que inserta los botones
                
                ' Crear cuadro combinado en D4 con los avisos
                Set ComboBox = wsSeleccionada.OLEObjects.Add(ClassType:="Forms.ComboBox.1", Link:=False, DisplayAsIcon:=False, _
                    Left:=wsSeleccionada.Range("D4").Left, Top:=wsSeleccionada.Range("D4").Top, _
                    Width:=wsSeleccionada.Range("D4").Width, Height:=wsSeleccionada.Range("D4").Height).Object
                ComboBox.List = Array("FAB", "AJ MEC", "CAL", "MPS", "MTO")
                ComboBox.Text = "AVISO"
                ComboBox.ForeColor = RGB(0, 0, 0)
                
                Exit For
            End If
        Next line
        
        If Not found Then
            MsgBox "Pieza no encontrada en el archivo CORDONES.txt.", vbExclamation
        End If
        
        ' Ajustar el ancho de la columna D al contenido
        wsSeleccionada.columns("D").AutoFit
        
        ' Combinar celdas de E1 a H1 y poner "OBSERVACIONES:"
        With wsSeleccionada.Range("E1:H1")
            .Merge
            .Value = "OBSERVACIONES:"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 14
            .Font.Bold = True
            .Interior.color = RGB(200, 200, 255)
        End With
        
        ' Combinar celdas de E2 a H6 y dejarlas en blanco con un tono de gris muy claro
        With wsSeleccionada.Range("E2:H6")
            .Merge
            .Value = ""
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .Font.Size = 14
            .Font.Bold = False
            .Interior.color = RGB(242, 242, 242) ' Tono de gris muy claro
        End With
        
        nombreLibro.GuardarNombreLibro
        nombreLibro.guardarOrigen
        ruta.AsignarRuta
        
        wsSeleccionada.Activate
        
    End If
    control.programarCierre newWorkbook
    MsgBox "Dispone de" & " " & control.tiempoCierre & " " & " minutos para cubrir el croquis. Pasado ese tiempo el croquis se cerrará sin guardar. Recuerde guardar su progreso.", vbExclamation
CleanUp:
    ' Restaurar la actualización de pantalla y el cálculo automático
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub



Sub CrearHojaConImagen(newWorkbook As Workbook, nombreHoja As String, rutaImagen As String)
    Dim ws As Worksheet
    Dim img As Picture
    
    ' Verificar si la hoja ya existe en el nuevo libro
    On Error Resume Next
    Set ws = newWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    
    ' Si la hoja no existe, crearla
    If ws Is Nothing Then
        Set ws = newWorkbook.Sheets.Add(After:=newWorkbook.Sheets(newWorkbook.Sheets.count))
        ws.name = nombreHoja
    End If
    
    ' Configurar la hoja sin encabezados, sin líneas de cuadrícula y sin barra de fórmulas
    With ws
        .Activate
        .Parent.Windows(1).DisplayGridlines = False
        .Parent.Windows(1).DisplayHeadings = False
        Application.DisplayFormulaBar = False
    End With
    
    ' Insertar la imagen en la hoja
    Set img = ws.Pictures.Insert(rutaImagen)
    img.Top = 60
    
    ' Girar la imagen 270 grados
    img.ShapeRange.Rotation = 270
    img.Height = 720
    img.Width = 480
    
End Sub


Function LimpiarNombreHoja(nombre As String) As String
    Dim caracteresInvalidos As String
    Dim i As Integer
    Dim caracter As String
    
    ' Definir caracteres inválidos para nombres de hojas
    caracteresInvalidos = "/\*[]:?<>|"
    
    ' Reemplazar caracteres inválidos con un guion bajo
    For i = 1 To Len(caracteresInvalidos)
        caracter = Mid(caracteresInvalidos, i, 1)
        nombre = Replace(nombre, caracter, "_")
    Next i
    
    ' Limitar el nombre a 31 caracteres (máximo permitido por Excel)
    If Len(nombre) > 31 Then
        nombre = Left(nombre, 31)
    End If
    
    LimpiarNombreHoja = nombre
End Function

Sub InsertarBotones(ws As Worksheet)
    Dim btnGuardar As Button
    Dim btnCancelar As Button
    
    ' Insertar botón "GUARDAR" en la celda A1
    Set btnGuardar = ws.Buttons.Add(ws.Cells(1, 1).Left, ws.Cells(1, 1).Top, ws.Cells(1, 1).Width, ws.Cells(1, 1).Height)
    With btnGuardar
        .Caption = "GUARDAR"
        .OnAction = "Guardar.Guardar" 'asigna el modulo guardar al boton
    End With
    
    ' Insertar botón "CANCELAR" en la celda A2
    Set btnCancelar = ws.Buttons.Add(ws.Cells(2, 1).Left, ws.Cells(2, 1).Top, ws.Cells(2, 1).Width, ws.Cells(2, 1).Height)
    With btnCancelar
        .Caption = "CANCELAR"
        .OnAction = "Cancelar.Cancelar" 'asigna el modulo cancelar al boton
    End With
End Sub



Private Sub CommandButton3_Click() '--BOTON FILTRAR
    Dim wsHome As Worksheet
    Dim wsRefs As Worksheet
    Dim ComboBox1 As Object
    Dim ComboBox3 As Object
    Dim robot As String
    Dim grupo As String
    Dim lastRow As Long
    Dim i As Long
    Dim found As Boolean
    
    'Referencia las hojas
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    Set wsRefs = ThisWorkbook.Sheets("referenciasHoy")
    
    ' Obtener el valor de la celda F4 de la hoja TECNICO
    robot = Trim(wsHome.Range("F4").Value)
    
    ' Referencia al ComboBox ActiveX en la hoja Home
    Set ComboBox1 = wsHome.OLEObjects("ComboBox1").Object
    Set ComboBox3 = wsHome.OLEObjects("ComboBox3").Object
    
    ' Obtener el valor seleccionado del ComboBox3 (grupos)
    grupo = Trim(ComboBox3.Value)
    
    ' Limpia el ComboBox antes de llenarlo
    ComboBox1.Clear
    
    ' Encuentra la última fila con datos en la columna E de la hoja referenciasHoy
    lastRow = wsRefs.Cells(wsRefs.Rows.count, "E").End(xlUp).Row
    
    ' Buscar el robot y/o grupo en las columnas E y A respectivamente y llenar el ComboBox con los valores de la columna B
    found = False
    For i = 2 To lastRow 'LOS DATOS EMPIEZAN EN LA FILA 2, LA FILA 1 SON LOS ENCABEZADOS
        If (robot = "" Or wsRefs.Cells(i, 5).Value = robot) And (grupo = "" Or wsRefs.Cells(i, 1).Value = grupo) Then
            ComboBox1.AddItem wsRefs.Cells(i, 2).Value
            found = True
        End If
    Next i
    
    ' Si no se encuentra ningún croquis, mostrar un mensaje
    If Not found Then
        MsgBox "No se encontraron croquis con los criterios especificados.", vbExclamation
    End If
End Sub


Private Sub CommandButton4_Click()  '--BOTON PREVISUALIZAR
    ' Subrutina previsualización
    Dim i As Integer
    Dim pieza As String
    Dim imagePath As String
    
    Dim wsHome As Worksheet
    Dim wsRefs As Worksheet
    Dim ComboBox As Object
    
    Set wsRefs = ThisWorkbook.Sheets("referenciasHoy")
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    
    Set ComboBox = wsHome.OLEObjects("ComboBox1").Object
    
    If ComboBox.Value = "" Then
        MsgBox "Debe indicar una pieza.", vbExclamation
        Exit Sub
    End If
    
    pieza = ComboBox.Value
    i = 0
    
    ' Buscar la pieza en la hoja de referencias
    Do While wsRefs.Range("B2").Offset(i, 0).Value <> pieza
        i = i + 1
        ' Evitar bucle infinito en caso de que no se encuentre
        If wsRefs.Range("B2").Offset(i, 0).Value = "" Then
            MsgBox "Pieza no encontrada.", vbExclamation
            Exit Sub
        End If
    Loop
    ' si la encuentra se queda con la ruta
    imagePath = wsRefs.Range("C2").Offset(i, 0).Value
    
    ' Verificar si la ruta de la imagen es válida
    If Dir(imagePath) = "" Then
        MsgBox "La ruta de la imagen no es válida.", vbExclamation
        Exit Sub
    End If
    
    ' Mostrar la imagen en el UserForm
    With UserForm1.Image1
        .Picture = LoadPicture(imagePath)
        .PictureSizeMode = fmPictureSizeModeStretch
    End With
    
    ' Mostrar el UserForm
    UserForm1.Show
End Sub

Private Sub CommandButton5_Click()  '--BOTÓN MODIFICAR
    Dim sheetName As String
    Dim wsHome As Worksheet
    Dim userSheet As Worksheet
    Dim foundUser As Range
    Dim libro As Workbook
    Dim rutaArchivo As String
    Dim btnGuardar As Button
    Dim btnCancelar As Button
    Dim fechaCarpeta As String
    Dim carpetaFecha As String
    Dim fileName As String
    Dim fileList As Collection
    Dim i As Integer
    Dim userName As String
    Dim checkBoxCount As Integer
    Dim mesas(2) As String
    Dim lastRow As Long
    Dim oleObj As OLEObject
    Dim oleObjComboBox As OLEObject
    Dim mesaComboBox As String
    
    mesas(0) = "M1"
    mesas(1) = "M2"
    mesas(2) = "M1/M2"
    
    ' Optimizar la ejecución desactivando algunas características de Excel
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' Asigna las hojas correspondientes
    Set wsHome = ThisWorkbook.Sheets("TECNICO")
    Set userSheet = ThisWorkbook.Sheets("Usuarios")
    
    ' Verificar si el usuario en G14 está en la columna A de la hoja Usuarios
    If Trim(wsHome.Range("G14").Value) = "" Then
        UserForm2.Show
        userName = Trim(userName) 'coge el username del UserForm2
        If userName = "" Then
            MsgBox "No se ingresó un nombre de usuario. Acceso denegado.", vbCritical
            ' Restaurar las configuraciones de Excel
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.DisplayAlerts = True
            Exit Sub
        End If
    Else
        userName = Trim(wsHome.Range("G14").Value) 'si lo encuentra, lo guarda
    End If
    
    ' Buscar el usuario en la hoja de usuarios
    Set foundUser = userSheet.columns("A").Find(What:=userName, LookIn:=xlValues, LookAt:=xlWhole)
    If foundUser Is Nothing Then
        MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
        ' Restaurar las configuraciones de Excel
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Exit Sub
    Else
        wsHome.Range("G14").Value = userName
    End If
    
    ' Extraer y formatear la fecha de la celda A14
    fechaCarpeta = Format(CDate(wsHome.Range("A14").Value), "yyyy_mm_dd")
    
    ' Definir la ruta de la carpeta con la fecha
    'carpetaFecha = ThisWorkbook.Path & "\Results\" & fechaCarpeta
    ruta.AsignarRuta
    carpetaFecha = ruta.ruta & "\" & fechaCarpeta & "\"
    
    
    ' Verificar si la carpeta con la fecha existe
    If Dir(carpetaFecha, vbDirectory) = "" Then
        MsgBox "No se encontraron documentos para la fecha especificada.", vbExclamation
        ' Restaurar las configuraciones de Excel
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Exit Sub 'y se sale de la subrutina
    End If
    
    ' Obtener la lista de archivos en la carpeta
    Set fileList = New Collection
    fileName = Dir(carpetaFecha & "\*.xlsx")
    Do While fileName <> ""
        fileList.Add fileName
        fileName = Dir
    Loop
    
    ' Verificar si hay archivos en la carpeta
    If fileList.count = 0 Then
        MsgBox "No se encontraron documentos para la fecha especificada.", vbExclamation
        ' Restaurar las configuraciones de Excel
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Exit Sub ' y salir de la subrutina
    End If
    
    ' Verifica si el ComboBox está vacío
    sheetName = ComboBox2.Value
    If sheetName = "" Then
        MsgBox "Por favor, seleccione un libro.", vbExclamation
        ' Restaurar las configuraciones de Excel
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Exit Sub
    End If
    
    ' Definir la ruta del archivo en la carpeta con la fecha
    rutaArchivo = carpetaFecha & "\" & sheetName
    
    
    Dim blockMod As String
    
    'crea un archivo que impide guardar
    blockMod = ruta.ruta & "\lock.txt"
        
    If Dir(blockMod) <> "" Then
        MsgBox "El archivo está siendo modificado por otro usuario. Por favor, inténtelo más tarde."
        Exit Sub
    End If
    ' Create lock file
    Open blockMod For Output As #1
    Print #1, "Locked"
    Close #1
    ' Abrir el libro seleccionado
    Set libro = Workbooks.Open(rutaArchivo)
    
    ' Asignar nuevas macros a los botones "Guardar" y "Cancelar"
    On Error Resume Next
    Set btnGuardar = libro.Sheets(1).Buttons(1)
    Set btnCancelar = libro.Sheets(1).Buttons(2)
    On Error GoTo 0
    
    If Not btnGuardar Is Nothing Then
        btnGuardar.OnAction = "'" & ThisWorkbook.name & "'!GuardarCancelarMod.GuardarResultados"
    End If
    
    If Not btnCancelar Is Nothing Then
        btnCancelar.OnAction = "'" & ThisWorkbook.name & "'!GuardarCancelarMod.CancelarModificaciones"
    End If
    
    ' Manejar los checkboxes y comboboxes en la hoja de usuarios
    For j = 9 To 11
        checkBoxCount = 0
        For Each oleObj In libro.Sheets(1).OLEObjects
            If TypeName(oleObj.Object) = "CheckBox" Then
                If oleObj.TopLeftCell.Column = j Then
                    checkBoxCount = checkBoxCount + 1
                End If
            End If
        Next oleObj
        lastRow = checkBoxCount + 1 ' +1 para incluir la fila del título
        
        For i = 2 To lastRow
            ' Verificar si el checkbox está marcado
            For Each oleObj In libro.Sheets(1).OLEObjects
                If TypeName(oleObj.Object) = "CheckBox" Then
                    If oleObj.TopLeftCell.Row = i And oleObj.TopLeftCell.Column = j Then
                        ' Obtener el primer ComboBox asociado
                        Dim comboBoxFound As Boolean
                        comboBoxFound = False
                        For Each oleObjComboBox In libro.Sheets(1).OLEObjects
                            If TypeName(oleObjComboBox.Object) = "ComboBox" Then
                                If oleObjComboBox.TopLeftCell.Row = i And oleObjComboBox.TopLeftCell.Column = j Then
                                    If Not comboBoxFound Then
                                        comboBoxFound = True
                                        mesaComboBox = oleObjComboBox.Object.Value
                                        If Trim(mesaComboBox) = "Mesa" Then
                                            With oleObjComboBox.Object
                                                .List = mesas
                                                .Text = "Mesa"
                                                .ForeColor = RGB(0, 0, 0)
                                                .ColumnCount = 1
                                                .ColumnWidths = "150 pt" ' Ajusta el ancho de la columna para que el contenido se vea completo
                                            End With
                                        Else
                                            With oleObjComboBox.Object
                                                .List = mesas
                                                .Text = mesaComboBox
                                                .ForeColor = RGB(0, 0, 0)
                                                .ColumnCount = 1
                                                .ColumnWidths = "150 pt" ' Ajusta el ancho de la columna para que el contenido se vea completo
                                            End With
                                        End If
                                    End If
                                End If
                            End If
                        Next oleObjComboBox
                    End If
                End If
            Next oleObj
        Next i
    Next j
    control.programarCierre libro
    ' Restaurar las configuraciones de Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "Dispone de" & " " & control.tiempoCierre & " " & " minutos para cubrir el croquis. Pasado ese tiempo el croquis se cerrará sin guardar. Recuerde guardar su progreso.", vbExclamation
End Sub



Private Sub CommandButton6_Click() 'BOTON CERRAR SESIÓN

    'REFERENCIA LA HOJA INICIO
    Dim homeSheetName As String
    homeSheetName = "INICIO"
    
    ' Verifica si la hoja Home existe
    On Error Resume Next
    Set wsHome = ThisWorkbook.Sheets(homeSheetName)
    On Error GoTo 0
    
    If Not wsHome Is Nothing Then
        ' Desoculta la hoja si está oculta
        wsHome.Visible = xlSheetVisible
        ' Activa la hoja Home
        wsHome.Activate
        ' Oculta la hoja DE TECNICOS
        Me.Visible = xlSheetVeryHidden
    Else
        MsgBox "La hoja " & homeSheetName & " no existe.", vbExclamation
    End If
End Sub

