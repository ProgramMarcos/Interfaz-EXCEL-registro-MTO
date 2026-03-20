VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    HOJA INFORMES
'------------------------------------------------------------------
'------------------------------------------------------------------

Private Sub CommandButton1_Click() '-- BOTON RESULTADOS
    Dim sheetName As String
    Dim ws As Worksheet
    Dim wsHome As Worksheet
    Dim userSheet As Worksheet
    Dim foundUser As Range
    
    
    'referenciado de las hojas
    sheetName = "Resultados"
    
    
    
    Dim libroOrigen As Workbook
    Dim carpetaResultados As String

    carpetaResultados = ruta.ruta

    ' Referencia al libro de origen usando la variable nombreLibro
    nombreLibro.GuardarNombreLibro
    Set libroOrigen = Workbooks(nombreLibro.nombreLibro)
    
    
    
    Set userSheet = ThisWorkbook.Sheets("Informantes")
    
    Set wsHome = libroOrigen.Sheets("INFORMES")
    
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
    Set userSheet = libroOrigen.Sheets("Informantes")
    Set foundUser = userSheet.columns("A").Find(What:=Trim(userName), LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundUser Is Nothing Then
        MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Desactivar la actualización de pantalla
    Application.ScreenUpdating = False
    On Error Resume Next
    ' Intenta abrir el libro de resultados
    Set libroResultados = Workbooks.Open(carpetaResultados & "\Resultados.xlsm")
    ' Verifica si el libro se abrió correctamente
    If libroResultados Is Nothing Then
        MsgBox "No se pudo abrir el libro de resultados.", vbCritical
        Exit Sub
    End If
    
    ' Intenta establecer la referencia a la hoja "Resultados"
    Set wsResultados = libroResultados.Sheets("Resultados")
    ' Verifica si la hoja "Resultados" existe
    If wsResultados Is Nothing Then
        MsgBox "La hoja 'Resultados' no existe en el libro de resultados.", vbCritical
        libroResultados.Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo 0
    
    
    ' Reactivar la actualización de pantalla
    Application.ScreenUpdating = True
    
    
End Sub

Private Sub CommandButton2_Click() '---ACCESO A TECNICO
    Dim sheetName As String
    Dim wsInicio As Worksheet
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim userSheet As Worksheet
    Dim foundUser As Range
    
    Set wsInicio = ThisWorkbook.Sheets("INFORMES")
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
    MsgBox "Se produjo un error: " & Err.Description, vbCritical
End Sub
Private Sub CommandButton3_Click() ' BOTÓN DE EXPORTAR PDF
    Dim ws As Worksheet
    Dim wsHome As Worksheet
    Dim userSheet As Worksheet
    Dim foundUser As Range
    Dim folderPath As String
    Dim savePath As Variant
    Dim fechaCarpeta As String
    Dim carpetaFecha As String
    Dim fileName As String
    Dim fileList As Collection
    Dim i As Integer
    Dim selectedFiles As Variant
    Dim pdfPath As String
    Dim libro As Workbook
    Dim fd As FileDialog
    Dim fileChosen As Boolean
    
    Set wsHome = ThisWorkbook.Sheets("INFORMES")
    Set userSheet = ThisWorkbook.Sheets("Informantes")
    
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
    Set foundUser = userSheet.columns("A").Find(What:=Trim(userName), LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundUser Is Nothing Then
        MsgBox "Usuario no autorizado. Acceso denegado.", vbCritical
        Exit Sub
    End If
    
    ' Obtener la fecha de la celda A14 y formatearla
    fechaCarpeta = Format(CDate(ThisWorkbook.Sheets("INFORMES").Range("A14").Value), "yyyy_mm_dd")
    
    ' Definir la ruta de la carpeta con la fecha
    ruta.AsignarRuta
    carpetaFecha = ruta.ruta & "\Results\" & fechaCarpeta
    
    ' Verificar si la carpeta con la fecha existe
    If Dir(carpetaFecha, vbDirectory) = "" Then
        MsgBox "No se encontraron documentos para la fecha especificada.", vbExclamation
        Exit Sub
    End If
    
    ' Crear un FileDialog para seleccionar múltiples archivos
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = "Seleccione los archivos de Excel para convertir a PDF"
        .InitialFileName = carpetaFecha & "\"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx"
        .AllowMultiSelect = True
        fileChosen = .Show
    End With
    
    ' Verificar si el usuario canceló el cuadro de diálogo
    If Not fileChosen Then
        MsgBox "Operación cancelada por el usuario.", vbExclamation
        Exit Sub
    End If
    
    ' Mostrar mensaje de inicio
    MsgBox "Este proceso tardará un rato, sea paciente", vbInformation
    
    ' Desactivar la actualización de pantalla y el cálculo automático para mejorar el rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' Convertir cada archivo seleccionado a PDF
    For i = 1 To fd.SelectedItems.count
        ' Abrir el libro de Excel
        Set libro = Workbooks.Open(fd.SelectedItems(i))
        
        ' Configurar la hoja para guardar en tamaño A4 y en formato horizontal
        For Each ws In libro.Worksheets
            With ws.PageSetup
                .PaperSize = xlPaperA4
                .Orientation = xlLandscape
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
            End With
        Next ws
        
        ' Definir la ruta para guardar el PDF
        pdfPath = Replace(fd.SelectedItems(i), ".xlsx", ".pdf")
        
        ' Guardar el libro como PDF
        libro.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        
        ' Cerrar el libro de Excel
        libro.Close SaveChanges:=False
    Next i
    
    ' Restaurar las configuraciones de Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    ' Mostrar mensaje de finalización
    MsgBox "Las hojas se han guardado como PDF.", vbInformation
End Sub






Private Sub CommandButton4_Click() 'CERRAR SESIÓN


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
        ' Oculta la hoja DE Informes
        Me.Visible = xlSheetVeryHidden
    Else
        MsgBox "La hoja " & homeSheetName & " no existe.", vbExclamation
    End If
End Sub

