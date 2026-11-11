Imports System.Windows

Public Class SocioItem
    Public Property IdSocio As Integer
    Public Property NombreCompleto As String
    Public Property Dni As Long
    Public Property DniFormateado As String
    Public Property Email As String
    Public Property Telefono As String
    Public Property Estado As EstadoSocio

    Public ReadOnly Property EstadoTexto As String
        Get
            Return If(Estado = EstadoSocio.Activo, "Habilitado", "Inhabilitado")
        End Get
    End Property

    Public ReadOnly Property EsActivo As Boolean
        Get
            Return Estado = EstadoSocio.Activo
        End Get
    End Property

    Public ReadOnly Property ColorBadgeFondo As String
        Get
            Return If(Estado = EstadoSocio.Activo, "#E7F3EC", "#FDECE9")
        End Get
    End Property

    Public ReadOnly Property ColorBadgeTexto As String
        Get
            Return If(Estado = EstadoSocio.Activo, "#1D3B31", "#991B1B")
        End Get
    End Property

    Public ReadOnly Property VisibilidadBotonDesactivar As Visibility
        Get
            Return If(Estado = EstadoSocio.Activo, Visibility.Visible, Visibility.Collapsed)
        End Get
    End Property

    Public ReadOnly Property VisibilidadBotonActivar As Visibility
        Get
            Return If(Estado = EstadoSocio.Inactivo, Visibility.Visible, Visibility.Collapsed)
        End Get
    End Property
End Class
