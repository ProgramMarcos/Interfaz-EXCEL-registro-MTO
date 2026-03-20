Attribute VB_Name = "ActiveCheckboxes"
'------------------------------------------------------------------
'------------------------------------------------------------------
'---    MARCOS LÓPEZ LÓPEZ
'---    2024/2025
'---    MODULO ActiveCheckBoxes
'---    Se inicializan los checkboxes, usuarios y admins
'------------------------------------------------------------------
'------------------------------------------------------------------

' En un módulo estándar
Public ActiveCheckboxes As Object

Sub InitializeDictionary()
    Set ActiveCheckboxes = CreateObject("Scripting.Dictionary")
End Sub

' Usuario
Public userName As String
'Administrador de usuarios
Public Admin As String


