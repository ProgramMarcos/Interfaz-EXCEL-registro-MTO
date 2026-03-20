Attribute VB_Name = "Cancelar"
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    MODULO Cancelar
'---    Contiene la subrutina asignada al botón "Cancelar"
'---    cuando se crea una hoja de reporte
'------------------------------------------------------------------
'------------------------------------------------------------------

Sub Cancelar()
    Dim sheetName As String
    Dim ws As Worksheet
    Dim currentSheet As Worksheet
    Dim oleObj As OLEObject
    Dim checkBoxMarcado As Boolean
    Dim nuevoLibro As Workbook
    
    
    
    sheetName = "TECNICO"
    
    ' Referencia al libro actual (el libro nuevo)
    Set nuevoLibro = ActiveWorkbook
    
    ' Verifica si la hoja "TECNICO" existe en el libro de origen
    On Error Resume Next
    Set ws = Workbooks("LibroOrigen.xlsx").Sheets(sheetName) ' Reemplaza "LibroOrigen.xlsx" con el nombre del libro de origen
    On Error GoTo 0
    
    Dim blockMod As String
    
    'crea un archivo que impide guardar
    blockMod = ruta.ruta & "\lock.txt"
        
    If Dir(blockMod) <> "" Then
        ' borrar el bloqueo si existe
        Kill blockMod
    End If
    
    If Not ws Is Nothing Then
        ' Obtén la hoja actual
        Set currentSheet = nuevoLibro.ActiveSheet
        
        ' Verifica si hay algún checkbox marcado en la hoja actual
        checkBoxMarcado = False
        For Each oleObj In currentSheet.OLEObjects
            If TypeName(oleObj.Object) = "CheckBox" Then
                If oleObj.Object.Value = True Then
                    checkBoxMarcado = True
                    Exit For
                End If
            End If
        Next oleObj
        
        ' Si no hay ningún checkbox marcado, elimina la hoja actual
        If Not checkBoxMarcado Then
            Application.DisplayAlerts = False
            currentSheet.Delete
            Application.DisplayAlerts = True
        Else
            ' Oculta la hoja actual
            currentSheet.Visible = xlSheetHidden
        End If
        
        ' Desoculta la hoja "TECNICO" si está oculta
        ws.Visible = xlSheetVisible
        
        ' Activa la hoja "TECNICO"
        ws.Activate
    Else
       ' MsgBox "La hoja " & sheetName & " no existe.", vbExclamation
    End If
    ' Llama al procedimiento cancelarCierre en el módulo control
    control.cancelarCierre
    ' Cierra el libro nuevo sin guardar
    nuevoLibro.Close SaveChanges:=False
End Sub

