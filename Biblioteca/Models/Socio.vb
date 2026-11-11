Public Enum EstadoSocio
    Activo = 0
    Inactivo = 1
End Enum

Public Class Socio
    Public Property idSocio As Integer
    Public Property nombreCompleto As String
    Public Property dni As Long
    Public Property email As String
    Public Property telefono As String
    Public Property estado As EstadoSocio = EstadoSocio.Activo
End Class
