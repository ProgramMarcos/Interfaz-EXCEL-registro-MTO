Attribute VB_Name = "NombreLibro"
Public nombreLibro As String
Public rutaOrigen As String

Sub GuardarNombreLibro()
    nombreLibro = ThisWorkbook.name
End Sub

Sub guardarOrigen()
    rutaOrigen = ThisWorkbook.Path
End Sub
