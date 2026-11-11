Imports System.Globalization

Public Class PrestamoBLL

    Private ReadOnly _prestamoDAL As New PrestamoDAL()
    Private ReadOnly _socioDAL As New SocioDAL()
    Private ReadOnly _libroDAL As New LibroDAL()
    Private Shared ReadOnly _culturaEs As New CultureInfo("es-AR")

    ''' <summary>
    ''' Obtiene el listado de préstamos formateados para su visualización en la UI.
    ''' </summary>
    Public Function solicitarListaPrestamos(Optional filtro As String = "") As List(Of PrestamoItem)
        Dim prestamos = _prestamoDAL.buscarPrestamos(filtro)
        Dim listaItems As New List(Of PrestamoItem)()

        For Each p In prestamos
            ' Color según plazo: Verde (#15803D) si está en plazo, Rojo (#DC2626) si ya venció
            Dim esVencido As Boolean = (p.fechaLimite.Date < Date.Today)
            Dim colorFecha As String = If(esVencido, "#DC2626", "#15803D")

            Dim fechaDevStr As String = If(p.fechaDevolucion.HasValue, p.fechaDevolucion.Value.ToString("dd MMM yyyy", _culturaEs).Replace(".", "").ToLower(), String.Empty)

            listaItems.Add(New PrestamoItem() With {
                .IdPrestamo = p.idPrestamo,
                .IdSocio = p.socio.idSocio,
                .NombreSocio = p.socio.nombreCompleto,
                .DniSocio = p.socio.dni,
                .IdLibro = p.libro.idLibro,
                .TituloLibro = p.libro.titulo,
                .AutorLibro = p.libro.autor,
                .IsbnLibro = p.libro.isbn,
                .FechaPrestamo = p.fechaPrestamo,
                .FechaPrestamoFormateada = p.fechaPrestamo.ToString("dd MMM yyyy", _culturaEs).Replace(".", "").ToLower(),
                .FechaLimite = p.fechaLimite,
                .FechaLimiteFormateada = p.fechaLimite.ToString("dd MMM yyyy", _culturaEs).Replace(".", "").ToLower(),
                .FechaDevolucion = p.fechaDevolucion,
                .FechaDevolucionFormateada = fechaDevStr,
                .FueDevuelto = p.fueDevuelto,
                .ColorFechaLimite = colorFecha
            })
        Next

        Return listaItems
    End Function

    ''' <summary>
    ''' Registra un nuevo préstamo para el libro y socio indicados, calculando las fechas de inicio y límite.
    ''' </summary>
    ''' <param name="idLibro">Identificador del libro a prestar.</param>
    ''' <param name="idSocio">Identificador del socio solicitante.</param>
    ''' <param name="plazo">Plazo en días del préstamo (ej. 7, 14, 30).</param>
    Public Function registrarPrestamo(idLibro As Integer, idSocio As Integer, plazo As Integer) As Prestamo
        ' Obtener socio y libro
        Dim socio As Socio = _socioDAL.buscarSocioPorId(idSocio)
        If socio Is Nothing OrElse socio.estado <> EstadoSocio.Activo Then
            Throw New InvalidOperationException($"No se encontró el socio con ID {idSocio}.")
        End If

        Dim libro As Libro = _libroDAL.buscarLibros().FirstOrDefault(Function(l) l.idLibro = idLibro)
        If libro Is Nothing Then
            Throw New InvalidOperationException($"No se encontró el libro con ID {idLibro}.")
        End If

        If libro.stockDisponible <= 0 Then
            Throw New InvalidOperationException($"El libro '{libro.titulo}' no cuenta con ejemplares disponibles para préstamo.")
        End If

        ' Calcular fecha actual y fecha límite
        Dim fechaActual As Date = Date.Today
        Dim fechaLimite As Date = fechaActual.AddDays(plazo)

        ' Instanciar Prestamo con su socio y su libro
        Dim nuevoPrestamo As New Prestamo() With {
            .socio = socio,
            .libro = libro,
            .fechaPrestamo = fechaActual,
            .fechaLimite = fechaLimite,
            .fechaDevolucion = Nothing,
            .fueDevuelto = False
        }

        ' Persistir objeto en la DB mediante PrestamoDAL
        _prestamoDAL.guardarPrestamo(nuevoPrestamo)

        Return nuevoPrestamo
    End Function

    ''' <summary>
    ''' Registra la devolución de un préstamo y reintegra el ejemplar al stock disponible.
    ''' </summary>
    Public Function registrarDevolucion(idPrestamo As Integer) As Boolean
        Return _prestamoDAL.registrarDevolucion(idPrestamo)
    End Function

End Class
