Public Class PrestamoItem
    Public Property IdPrestamo As Integer
    Public Property IdSocio As Integer
    Public Property NombreSocio As String
    Public Property DniSocio As Long
    Public Property IdLibro As Integer
    Public Property TituloLibro As String
    Public Property AutorLibro As String
    Public Property IsbnLibro As Long
    Public Property FechaPrestamo As Date
    Public Property FechaPrestamoFormateada As String
    Public Property FechaLimite As Date
    Public Property FechaLimiteFormateada As String
    Public Property FechaDevolucion As Date?
    Public Property FechaDevolucionFormateada As String
    Public Property FueDevuelto As Boolean

    ''' <summary>
    ''' Color del texto para la fecha límite:
    ''' Verde (#15803D) si está dentro del plazo, Rojo (#DC2626) si ya venció el plazo.
    ''' </summary>
    Public Property ColorFechaLimite As String

    Public ReadOnly Property VisibilidadBotonDevolver As Visibility
        Get
            Return If(FueDevuelto, Visibility.Collapsed, Visibility.Visible)
        End Get
    End Property

    Public ReadOnly Property VisibilidadTextoDevuelto As Visibility
        Get
            Return If(FueDevuelto, Visibility.Visible, Visibility.Collapsed)
        End Get
    End Property
End Class
