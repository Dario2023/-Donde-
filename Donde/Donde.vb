
Public Enum TipoAccion As Byte 'Enumeracion para selecionar el tipo de accion a realizar por el usuario
    Alta = 1
    Modificacion = 2
    Baja = 3
End Enum


Public Class F_Donde

    Friend Accion As TipoAccion
    Friend IdPersona As Long

    Private Sub Cargar_Persona(sender As Object, e As EventArgs) Handles Btn_AP.Click ' Llama al formolario para realizar la carga de los datos de una perosna a la Base de Datos
        IdPersona = -1
        Accion = TipoAccion.Alta
        F_CPersona.Show()
    End Sub








End Class
