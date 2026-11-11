Public Class Prestamo
    Public Property idPrestamo As Integer
    Public Property socio As Socio
    Public Property libro As Libro
    Public Property fechaPrestamo As Date
    Public Property fechaLimite As Date
    Public Property fechaDevolucion As Date?
    Public Property fueDevuelto As Boolean
End Class
