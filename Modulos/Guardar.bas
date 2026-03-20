Attribute VB_Name = "Guardar"
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    MODULO Guardar
'---    Contiene la subrutina asignada al botón "Guardar"
'---    cuando se crea una hoja de reporte
'------------------------------------------------------------------
'------------------------------------------------------------------
Sub Guardar()
    Dim ws As Worksheet
    Dim wsHome As Worksheet
    Dim wsResultados As Worksheet
    Dim wsReferenciasHoy As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long, k As Long
    Dim oleObj As OLEObject
    Dim oleObjComboBox1 As OLEObject
    Dim oleObjComboBox2 As OLEObject
    Dim titulo As String
    Dim nombreCheckbox As String
    Dim filaResultados As Long
    Dim ComboBox As Object
    Dim valorD1 As String
    Dim valorE1 As String
    Dim valorF2 As String
    Dim valorF1 As String
    Dim valorC1 As String
    Dim valorC2 As String
    Dim valorC3 As String
    Dim valorB1 As String
    Dim valorB2 As String
    Dim valorB3 As String
    Dim valorB4 As String
    Dim valorB4_replace As String
    Dim valorE2 As String
    Dim busqueda As String
    Dim zonaPieza As String
    Dim valorComboBox As String
    Dim existe As Boolean
    Dim itemExists As Boolean
    Dim itemIndex As Long
    Dim checkBoxCount As Long
    Dim valorComboBox1 As String
    Dim valorComboBox2 As String
    Dim valorComboBox3 As String
    Dim mesaComboBox As String
    Dim penultimoComboBox As OLEObject
    Dim ultimoComboBox As OLEObject
    Dim cnt As Integer
    Dim dadDag As String
    Dim valorColumnaM As String
    Dim sheetName As String
    Dim nuevoNombreHoja As String
    Dim nombreHoja As String
    Dim carpetaResultados As String
    Dim nuevoLibro As Workbook
    Dim rutaArchivo As String
    Dim fechaCarpeta As String
    Dim carpetaFecha As String
    Dim libroOrigen As Workbook
    
    Dim fecha As Date
    Dim hora As Date
    
    Dim libroResultados As Workbook
    Dim blockMod As String
    
    'crea un archivo que impide guardar
    blockMod = ruta.ruta & "\lock.txt"
'
'    If Dir(blockMod) <> "" Then
'        MsgBox "El archivo está siendo modificado por otro usuario. Por favor, inténtelo más tarde."
'        Exit Sub
'    End If
'    ' Create lock file
'    Open blockMod For Output As #1
'    Print #1, "Locked"
'    Close #1
    
    ' Referencia al libro actual (el libro nuevo)
    Set nuevoLibro = ActiveWorkbook
    
    ' Referencia a la hoja activa
    Set ws = nuevoLibro.ActiveSheet
    
    
    ' Obtener los valores de las celdas
    cnt = ws.Cells(3, 1).Value 'GUARDO EL NUMERO DE REGULACIÓN
    valorB1 = ws.Cells(1, 2).Value ' Celda B1 R2/BAV/XFK/VS20
    valorB2 = ws.Cells(2, 2).Value ' Celda B2 MAG/RP
    valorB3 = ws.Cells(3, 2).Value ' Celda B3 FECHA
    valorB4 = Format(ws.Cells(4, 2).Value, "HH:MM:SS") ' Celda B4 HORA (formato HH:MM:SS)
    valorC1 = ws.Cells(1, 3).Value ' Celda C1 SO-xxx
    valorC2 = ws.Cells(2, 3).Value ' Celda C2 ROBOT
    valorC3 = ws.Cells(3, 3).Value ' Celda C3 USER
    valorD1 = ws.Cells(1, 4).Value 'TITULO
    valorE2 = ws.Cells(2, 5).Value ' OBSERVACIONES (celda combinada E2:H6)
    
    ' Identificar el penúltimo y el último ComboBox
    Dim comboBoxList As Collection
    Set comboBoxList = New Collection
    For Each oleObj In ws.OLEObjects
        If TypeName(oleObj.Object) = "ComboBox" Then
            comboBoxList.Add oleObj
        End If
    Next oleObj
    
    If comboBoxList.count < 2 Then
        MsgBox "No hay suficientes ComboBoxes en la hoja.", vbExclamation
        Exit Sub
    End If
    
    Set penultimoComboBox = comboBoxList(comboBoxList.count - 1)
    Set ultimoComboBox = comboBoxList(comboBoxList.count)
    
    ' Verificar si el penúltimo ComboBox (mesa) está vacío
    If penultimoComboBox.Object.Value = "" Or Trim(penultimoComboBox.Object.Value) = "MESA" Then
        MsgBox "Debe seleccionar una mesa antes de guardar.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si SE HA INTRODUCIDO UNA HORA
    If ws.Cells(4, 2).Value = "" Then
        MsgBox "Debe introducir la hora de la regulación antes de guardar. Compruebe el cuadro gris en la parte superior izquierda.", vbExclamation
        Exit Sub
    End If
    
    ' Cambiar el nombre de la hoja
    valorB4_replace = Replace(valorB4, ":", "_") ' Reemplazar los dos puntos por guiones bajos
    nuevoNombreHoja = cnt & "_" & valorB4_replace & " " & valorC2 & " " & ws.name
    
    ' Verificar si el nombre excede los 31 caracteres y recortar si es necesario
    If Len(nuevoNombreHoja) > 31 Then
        nuevoNombreHoja = Left(nuevoNombreHoja, 31)
    End If
    
    ws.name = nuevoNombreHoja
    
    ' Formatear la fecha de la celda B3
    fechaCarpeta = Format(CDate(valorB3), "yyyy_mm_dd")
    
    ' Verificar si la carpeta "Results" existe, si no, crearla
    ruta.AsignarRuta
    carpetaResultados = ruta.ruta
    
    If Dir(carpetaResultados, vbDirectory) = "" Then
        MkDir carpetaResultados
    End If
    
    ' Verificar si la carpeta con la fecha existe, si no, crearla
    carpetaFecha = carpetaResultados & "\" & fechaCarpeta
    If Dir(carpetaFecha, vbDirectory) = "" Then
        MkDir carpetaFecha
    End If
    
    ' Guardar el libro actual en la carpeta con la fecha
    rutaArchivo = carpetaFecha & "\" & nuevoNombreHoja & ".xlsx"
    nuevoLibro.SaveAs rutaArchivo
    
    ' Referencia al libro de origen usando la variable nombreLibro
    nombreLibro.GuardarNombreLibro
    Set libroOrigen = Workbooks(nombreLibro.nombreLibro)
    
    ' Desactivar la actualización de pantalla
    Application.ScreenUpdating = False
    
    ' Referencia a la hoja "TECNICO" en el libro de origen
    Set wsHome = libroOrigen.Sheets("TECNICO")
    
    
    rutaLibro = carpetaResultados & "\Resultados.xlsm"
    
    ' Intenta abrir el libro de resultados
    On Error Resume Next
    Set libroResultados = Workbooks.Open(rutaLibro)
    On Error GoTo 0
    
    ' Si no se pudo abrir, crea un nuevo libro
    If libroResultados Is Nothing Then
        Set libroResultados = Workbooks.Add
        libroResultados.SaveAs fileName:=rutaLibro, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    End If
    
    ' Intenta establecer la referencia a la hoja "Resultados"
    On Error Resume Next
    Set wsResultados = libroResultados.Sheets("Resultados")
    On Error GoTo 0
    
    ' Si la hoja no existe, créala
    If wsResultados Is Nothing Then
        Set wsResultados = libroResultados.Sheets.Add(After:=libroResultados.Sheets(libroResultados.Sheets.count))
        wsResultados.name = "Resultados"
    End If

    On Error GoTo 0
    
    
    Set wsReferenciasHoy = libroOrigen.Sheets("referenciasHoy")
    
    '---------------------
    wsResultados.Cells(1, 1).Value = "FECHA" 'fecha
    wsResultados.Cells(1, 2).Value = "GRUPO REG" 'grupo reg
    wsResultados.Cells(1, 3).Value = "DETECCIÓN"
    wsResultados.Cells(1, 4).Value = "AVISO" ' Aviso
    wsResultados.Cells(1, 5).Value = "TEAM LEADER" ' Aviso
    wsResultados.Cells(1, 6).Value = "HORA" 'hora
    wsResultados.Cells(1, 7).Value = "PROYECTO" 'proyecto
    wsResultados.Cells(1, 8).Value = "TIPO SOLDADURA" 'rp/mag
    wsResultados.Cells(1, 9).Value = "PIEZA" 'pieza
    wsResultados.Cells(1, 10).Value = "MODELO" ' Guardar DAD o DAG
    wsResultados.Cells(1, 11).Value = "PUESTO" 'puesto
    wsResultados.Cells(1, 12).Value = "ROBOT" ' Aviso
    wsResultados.Cells(1, 13).Value = "CORDON" 'cordon
    wsResultados.Cells(1, 14).Value = "MESA" ' Mesa
    wsResultados.Cells(1, 15).Value = "CAUSA" 'CAUSA
    wsResultados.Cells(1, 16).Value = "PROBLEMA" 'PROBLEMA
    wsResultados.Cells(1, 17).Value = "ACCIÓN" 'acción
    
    wsResultados.Cells(1, 18).Value = "QUIÉN" 'USER
    wsResultados.Cells(1, 19).Value = "BÚSQUEDA"
    
    wsResultados.Cells(1, 20).Value = "ZONA PIEZA"
    wsResultados.Cells(1, 21).Value = "COMENTARIOS" 'COMENTARIOS
    
    With wsResultados.Range("A1:U1")
        .Font.Bold = True
        .Interior.color = RGB(255, 255, 0) ' Amarillo
    End With
    
    ' Inicializar la fila de resultados
    filaResultados = wsResultados.Cells(wsResultados.Rows.count, 1).End(xlUp).Row + 1
    
    
    
    ' Obtener los títulos y los nombres de los checkboxes marcados
    For j = 9 To 11 ' Columnas I, J, K (9, 10, 11)
        titulo = ws.Cells(1, j).Value
        If titulo <> "" Then
            ' Determinar si el título contiene "DAD" o "DAG"
            If InStr(1, titulo, " DAD", vbTextCompare) > 0 Then
                dadDag = "DAD"
                titulo = Replace(titulo, " DAD", "", 1, -1, vbTextCompare)
            ElseIf InStr(1, titulo, " DAG", vbTextCompare) > 0 Then
                dadDag = "DAG"
                titulo = Replace(titulo, " DAG", "", 1, -1, vbTextCompare)
            Else
                dadDag = ""
            End If
            
            ' Contar el número de checkboxes en la columna j
            checkBoxCount = 0
            For Each oleObj In ws.OLEObjects
                If TypeName(oleObj.Object) = "CheckBox" Then
                    If oleObj.TopLeftCell.Column = j Then
                        checkBoxCount = checkBoxCount + 1
                    End If
                End If
            Next oleObj
            
            lastRow = checkBoxCount + 1 ' +1 para incluir la fila del título
            
            For i = 2 To lastRow
                ' Verificar si el checkbox está marcado
                For Each oleObj In ws.OLEObjects
                    If TypeName(oleObj.Object) = "CheckBox" Then
                        If oleObj.TopLeftCell.Row = i And oleObj.TopLeftCell.Column = j Then
                            If oleObj.Object.Value = True Then
                                nombreCheckbox = oleObj.Object.Caption
                                
                                ' Obtener los valores de los ComboBox asociados
                                valorComboBox1 = ""
                                valorComboBox2 = ""
                                valorComboBox3 = ""
                                mesaComboBox = ""
                                Dim comboBoxCount As Integer
                                comboBoxCount = 0
                                For Each oleObjComboBox In ws.OLEObjects
                                    If TypeName(oleObjComboBox.Object) = "ComboBox" Then
                                        If oleObjComboBox.TopLeftCell.Row = i And oleObjComboBox.TopLeftCell.Column = j Then
                                            comboBoxCount = comboBoxCount + 1
                                            If comboBoxCount = 1 Then
                                                mesaComboBox = oleObjComboBox.Object.Value
                                            ElseIf comboBoxCount = 2 Then
                                                valorComboBox1 = oleObjComboBox.Object.Value
                                            ElseIf comboBoxCount = 3 Then
                                                valorComboBox2 = oleObjComboBox.Object.Value
                                            ElseIf comboBoxCount = 4 Then
                                                valorComboBox3 = oleObjComboBox.Object.Value
                                            End If
                                        End If
                                    End If
                                Next oleObjComboBox
                                
                                ' Verificar si el dato ya existe en la hoja "Resultados"
                                existe = False
                                For k = 1 To wsResultados.Cells(wsResultados.Rows.count, 1).End(xlUp).Row
                                    If wsResultados.Cells(k, 1).Value = valorB3 And wsResultados.Cells(k, 6).Value = valorB4 And wsResultados.Cells(k, 7).Value = valorB1 _
                                    And wsResultados.Cells(k, 9).Value = titulo And wsResultados.Cells(k, 10).Value = dadDag And wsResultados.Cells(k, 11).Value = valorC1 _
                                    And wsResultados.Cells(k, 14).Value = mesaComboBox And wsResultados.Cells(k, 13).Value = nombreCheckbox Then
                                        existe = True
                                        Exit For
                                    End If
                                Next k
                                
                                ' Si no existe, agregarlo
                                If Not existe Then
                                    If Trim(mesaComboBox) = "M1/M2" Then
                                    
                                        
                                        hora = TimeValue(valorB4)
                                        fecha = CDate(valorB3) ' convierte el string a fecha
                                        
                                        If (hora >= TimeValue("00:00:00") And hora < TimeValue("06:00:00")) Then
                                            wsResultados.Cells(filaResultados, 1).Value = fecha - 1 ' resta un día
                                        Else
                                            wsResultados.Cells(filaResultados, 1).Value = fecha
                                        End If
                                    
                                        'wsResultados.Cells(filaResultados, 1).Value = valorB3 'fecha
                                        wsResultados.Cells(filaResultados, 2).Value = cnt 'grupo reg
                                        wsResultados.Cells(filaResultados, 3).Value = " "
                                        wsResultados.Cells(filaResultados, 4).Value = ultimoComboBox.Object.Value ' Aviso
                                        wsResultados.Cells(filaResultados, 5).Value = " "
                                        wsResultados.Cells(filaResultados, 6).Value = valorB4 'hora
                                        wsResultados.Cells(filaResultados, 7).Value = valorB1 'proyecto
                                        wsResultados.Cells(filaResultados, 8).Value = valorB2 'rp/mag
                                        wsResultados.Cells(filaResultados, 9).Value = titulo 'pieza
                                        wsResultados.Cells(filaResultados, 10).Value = dadDag ' Guardar DAD o DAG
                                        wsResultados.Cells(filaResultados, 11).Value = valorC1 'puesto
                                        wsResultados.Cells(filaResultados, 12).Value = valorC2 'ROBOT
                                        wsResultados.Cells(filaResultados, 13).Value = nombreCheckbox 'cordon
                                        wsResultados.Cells(filaResultados, 14).Value = "M1" ' Mesa
                                        wsResultados.Cells(filaResultados, 18).Value = valorC3 'USER
                                        busqueda = valorB1 & titulo & nombreCheckbox 'valor de busqueda
                                        wsResultados.Cells(filaResultados, 19).Value = busqueda
                                        wsResultados.Cells(filaResultados, 21).Value = valorE2 'COMENTARIOS
                                        
                                        If Trim(valorComboBox1) = "PROBLEMA" Then
                                            wsResultados.Cells(filaResultados, 16).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 16).Value = valorComboBox1
                                        End If
                                        If Trim(valorComboBox2) = "CAUSA" Then
                                            wsResultados.Cells(filaResultados, 15).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 15).Value = valorComboBox2
                                        End If
                                        If Trim(valorComboBox3) = "SOLUCIÓN" Then
                                            wsResultados.Cells(filaResultados, 17).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 17).Value = valorComboBox3
                                        End If
                                        
                                        ' Buscar el valor en la hoja "referenciasHoy"
                                        valorColumnaM = ""
                                        For k = 1 To wsReferenciasHoy.Cells(wsReferenciasHoy.Rows.count, 9).End(xlUp).Row
                                            If wsReferenciasHoy.Cells(k, 9).Value = busqueda Then
                                                valorColumnaM = wsReferenciasHoy.Cells(k, 13).Value ' Columna M es la 13
                                                Exit For
                                            End If
                                        Next k
                                        
                                        ' Asignar el valor de la columna M a la columna 16 de la hoja de resultados
                                        wsResultados.Cells(filaResultados, 20).Value = valorColumnaM
                                        
                                        filaResultados = filaResultados + 1
                                        
                                        hora = TimeValue(valorB4)
                                        fecha = CDate(valorB3) ' convierte el string a fecha
                                        
                                        If (hora >= TimeValue("00:00:00") And hora < TimeValue("06:00:00")) Then
                                            wsResultados.Cells(filaResultados, 1).Value = fecha - 1 ' resta un día
                                        Else
                                            wsResultados.Cells(filaResultados, 1).Value = fecha
                                        End If
                                    
                                        wsResultados.Cells(filaResultados, 2).Value = cnt 'grupo reg
                                        wsResultados.Cells(filaResultados, 3).Value = " "
                                        wsResultados.Cells(filaResultados, 4).Value = ultimoComboBox.Object.Value ' Aviso
                                        wsResultados.Cells(filaResultados, 5).Value = " "
                                        wsResultados.Cells(filaResultados, 6).Value = valorB4 'hora
                                        wsResultados.Cells(filaResultados, 7).Value = valorB1 'proyecto
                                        wsResultados.Cells(filaResultados, 8).Value = valorB2 'rp/mag
                                        wsResultados.Cells(filaResultados, 9).Value = titulo 'pieza
                                        wsResultados.Cells(filaResultados, 10).Value = dadDag ' Guardar DAD o DAG
                                        wsResultados.Cells(filaResultados, 11).Value = valorC1 'puesto
                                        wsResultados.Cells(filaResultados, 12).Value = valorC2 'ROBOT
                                        wsResultados.Cells(filaResultados, 13).Value = nombreCheckbox 'cordon
                                        wsResultados.Cells(filaResultados, 14).Value = "M2" ' Mesa
                                        wsResultados.Cells(filaResultados, 18).Value = valorC3 'USER
                                        busqueda = valorB1 & titulo & nombreCheckbox 'valor de busqueda
                                        wsResultados.Cells(filaResultados, 19).Value = busqueda
                                        wsResultados.Cells(filaResultados, 21).Value = valorE2 'COMENTARIOS
                                        
                                        If Trim(valorComboBox1) = "PROBLEMA" Then
                                            wsResultados.Cells(filaResultados, 16).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 16).Value = valorComboBox1
                                        End If
                                        If Trim(valorComboBox2) = "CAUSA" Then
                                            wsResultados.Cells(filaResultados, 15).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 15).Value = valorComboBox2
                                        End If
                                        If Trim(valorComboBox3) = "SOLUCIÓN" Then
                                            wsResultados.Cells(filaResultados, 17).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 17).Value = valorComboBox3
                                        End If
                                        
                                        wsResultados.Cells(filaResultados, 20).Value = valorColumnaM
                                        
                                    ElseIf Trim(mesaComboBox) <> "M1/M2" And Trim(mesaComboBox) <> "Mesa" Then
                                        hora = TimeValue(valorB4)
                                        fecha = CDate(valorB3) ' convierte el string a fecha
                                        
                                        If (hora >= TimeValue("00:00:00") And hora < TimeValue("06:00:00")) Then
                                            wsResultados.Cells(filaResultados, 1).Value = fecha - 1 ' resta un día
                                        Else
                                            wsResultados.Cells(filaResultados, 1).Value = fecha
                                        End If
                                        wsResultados.Cells(filaResultados, 2).Value = cnt 'grupo reg
                                        wsResultados.Cells(filaResultados, 3).Value = " "
                                        wsResultados.Cells(filaResultados, 4).Value = ultimoComboBox.Object.Value ' Aviso
                                        wsResultados.Cells(filaResultados, 5).Value = " "
                                        wsResultados.Cells(filaResultados, 6).Value = valorB4 'hora
                                        wsResultados.Cells(filaResultados, 7).Value = valorB1 'proyecto
                                        wsResultados.Cells(filaResultados, 8).Value = valorB2 'rp/mag
                                        wsResultados.Cells(filaResultados, 9).Value = titulo 'pieza
                                        wsResultados.Cells(filaResultados, 10).Value = dadDag ' Guardar DAD o DAG
                                        wsResultados.Cells(filaResultados, 11).Value = valorC1 'puesto
                                        wsResultados.Cells(filaResultados, 12).Value = valorC2 'ROBOT
                                        wsResultados.Cells(filaResultados, 13).Value = nombreCheckbox 'cordon
                                        wsResultados.Cells(filaResultados, 14).Value = mesaComboBox ' Mesa
                                        wsResultados.Cells(filaResultados, 18).Value = valorC3 'USER
                                        busqueda = valorB1 & titulo & nombreCheckbox 'valor de busqueda
                                        wsResultados.Cells(filaResultados, 19).Value = busqueda
                                        wsResultados.Cells(filaResultados, 21).Value = valorE2 'COMENTARIOS
                                        
                                        If Trim(valorComboBox1) = "PROBLEMA" Then
                                            wsResultados.Cells(filaResultados, 16).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 16).Value = valorComboBox1
                                        End If
                                        If Trim(valorComboBox2) = "CAUSA" Then
                                            wsResultados.Cells(filaResultados, 15).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 15).Value = valorComboBox2
                                        End If
                                        If Trim(valorComboBox3) = "SOLUCIÓN" Then
                                            wsResultados.Cells(filaResultados, 17).Value = " "
                                        Else
                                            wsResultados.Cells(filaResultados, 17).Value = valorComboBox3
                                        End If
                                        
                                        ' Buscar el valor en la hoja "referenciasHoy"
                                        valorColumnaM = ""
                                        For k = 1 To wsReferenciasHoy.Cells(wsReferenciasHoy.Rows.count, 9).End(xlUp).Row
                                            If wsReferenciasHoy.Cells(k, 9).Value = busqueda Then
                                                valorColumnaM = wsReferenciasHoy.Cells(k, 13).Value ' Columna M es la 13
                                                Exit For
                                            End If
                                        Next k
                                        
                                        ' Asignar el valor de la columna M a la columna 16 de la hoja de resultados
                                        wsResultados.Cells(filaResultados, 20).Value = valorColumnaM
                                        
                                    End If
                                    
                                    filaResultados = filaResultados + 1
                                End If
                                
                            Else
                                nombreCheckbox = oleObj.Object.Caption
                                
                                ' Obtener los valores de los ComboBox asociados
                                valorComboBox1 = ""
                                valorComboBox2 = ""
                                valorComboBox3 = ""
                                mesaComboBox = ""
                                comboBoxCount = 0
                                For Each oleObjComboBox In ws.OLEObjects
                                    If TypeName(oleObjComboBox.Object) = "ComboBox" Then
                                        If oleObjComboBox.TopLeftCell.Row = i And oleObjComboBox.TopLeftCell.Column = j Then
                                            comboBoxCount = comboBoxCount + 1
                                            If comboBoxCount = 1 Then
                                                oleObjComboBox.Object.Value = ""
                                            ElseIf comboBoxCount = 2 Then
                                                oleObjComboBox.Object.Value = ""
                                            ElseIf comboBoxCount = 3 Then
                                                oleObjComboBox.Object.Value = ""
                                            ElseIf comboBoxCount = 4 Then
                                                oleObjComboBox.Object.Value = ""
                                            End If
                                        End If
                                    End If
                                Next oleObjComboBox
                            
                            
                            End If
                        End If
                    End If
                Next oleObj
            Next i
        End If
    Next j
    
    ' Ajustar el tamaño de las columnas al contenido
    wsResultados.columns("A:U").AutoFit
    
    Dim color As Integer
    For color = 1 To 21 ' Desde la columna 1 (A) hasta la columna 21 (U)
        If color Mod 2 = 0 Then ' Columnas pares
            wsResultados.columns(color).Interior.color = RGB(255, 255, 255) ' Blanco
        Else ' Columnas impares
            wsResultados.columns(color).Interior.color = RGB(211, 211, 211) ' Gris claro
        End If
        
        ' Remarcar las cuadrículas
        With wsResultados.columns(color).Borders
            .LineStyle = xlContinuous
            .color = RGB(0, 0, 0) ' Color negro para las cuadrículas
            .TintAndShade = 0
            .Weight = xlThin
        End With
    Next color
    
    
    ' Cierra el libro guardando cambios
    libroResultados.Close SaveChanges:=True
    
    ' Reactivar la actualización de pantalla
    Application.ScreenUpdating = True
    
    '---------------------
    
    ' Incrementar el contador en la celda G6
    'wsHome.Range("G6").Value = wsHome.Range("G6").Value + 1
    
    ' borrar el bloqueo al terminar
    Kill blockMod
    ' Llama al procedimiento cancelarCierre en el módulo control
    control.cancelarCierre
    ' Cerrar el libro nuevo sin guardar
    nuevoLibro.Close SaveChanges:=True
End Sub

