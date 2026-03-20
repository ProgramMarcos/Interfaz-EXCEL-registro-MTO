Attribute VB_Name = "control"
' Módulo control

' Variable global para almacenar el nombre del libro
Public nombreLibro As String
Dim tiempoEjecucion As Date
Public tiempoCierre As Date
Dim cerrarProgramado As Boolean

Sub programarCierre(libro As Workbook)
    ' Establece el tiempo de ejecución para 5 minutos en el futuro
    tiempoCierre = TimeValue("00:05:00")
    tiempoEjecucion = Now + tiempoCierre
    
    ' Guarda el nombre del libro en una variable global
    nombreLibro = libro.name
    
    ' Establece la bandera para indicar que el cierre está programado
    cerrarProgramado = True
    
    ' Programa la ejecución del procedimiento cerrarLibro
    Application.OnTime EarliestTime:=tiempoEjecucion, Procedure:="cerrarLibro", Schedule:=True
End Sub

Sub cancelarCierre()
    ' Cancela la ejecución programada del procedimiento cerrarLibro
    cerrarProgramado = False
End Sub

Sub cerrarLibro()
    ' Verifica si el cierre aún está programado
    If cerrarProgramado Then
        ' Verifica si el libro aún está abierto
        Dim libroAbierto As Boolean
        libroAbierto = False
        Dim wb As Workbook
        For Each wb In Workbooks
            If wb.name = nombreLibro Then
                libroAbierto = True
                Exit For
            End If
        Next wb
        
        ' Si el libro está abierto, ciérralo
        If libroAbierto Then
            Workbooks(nombreLibro).Close SaveChanges:=False
        End If
        
        ' Restablece la bandera
        cerrarProgramado = False
    End If
End Sub


