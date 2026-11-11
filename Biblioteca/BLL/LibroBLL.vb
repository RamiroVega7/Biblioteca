Public Class LibroBLL

    Private ReadOnly _libroDAL As New LibroDAL()

    Public Function solicitarListaLibros(Optional filtro As String = "") As List(Of LibroItem)
        Dim libros As List(Of Libro) = _libroDAL.buscarLibros(filtro)
        Dim listaItems As New List(Of LibroItem)()

        For Each l In libros
            Dim tieneStock As Boolean = (l.stockDisponible > 0)
            Dim textoStock As String
            Dim colorStock As String
            Dim colorBadgeStock As String
            Dim colorBoton As String

            If l.stockDisponible <= 0 Then
                textoStock = "0 (Sin stock)"
                colorStock = "#DC2626"
                colorBadgeStock = "#FDECE9"
                colorBoton = "#9CA3AF"
            ElseIf l.stockDisponible = 1 Then
                textoStock = "1 ejemplar"
                colorStock = "#15803D"
                colorBadgeStock = "#E7F3EC"
                colorBoton = "#1D3B31"
            Else
                textoStock = $"{l.stockDisponible} ejemplares"
                colorStock = "#15803D"
                colorBadgeStock = "#E7F3EC"
                colorBoton = "#1D3B31"
            End If

            listaItems.Add(New LibroItem() With {
                .IdLibro = l.idLibro,
                .Isbn = l.isbn,
                .IsbnFormateado = FormatearIsbn(l.isbn),
                .Titulo = l.titulo,
                .Autor = l.autor,
                .Editorial = l.editorial,
                .StockTotal = l.stockTotal,
                .StockTotalTexto = $"{l.stockTotal} ejemplares",
                .StockDisponible = l.stockDisponible,
                .StockDisponibleTexto = textoStock,
                .StockRatioTexto = $"{l.stockDisponible} / {l.stockTotal}",
                .ColorStock = colorStock,
                .ColorBadgeStock = colorBadgeStock,
                .TieneStock = tieneStock,
                .ColorBotonSeleccionar = colorBoton
            })
        Next

        Return listaItems
    End Function

    Private Shared Function FormatearIsbn(isbn As Long) As String
        Dim s As String = isbn.ToString()
        If s.Length = 13 Then
            Select Case isbn
                Case 9788437604947L
                    Return "978-84-376-0494-7"
                Case 9788478887194L
                    Return "978-84-7888-719-4"
                Case 9788408043645L
                    Return "978-84-08-04364-5"
                Case 9788420467284L
                    Return "978-84-204-6728-4"
                Case 9788437604572L
                    Return "978-84-376-0457-2"
                Case 9788497592208L
                    Return "978-84-9759-220-8"
                Case 9788466319128L
                    Return "978-84-663-1912-8"
                Case Else
                    Return $"{s.Substring(0, 3)}-{s.Substring(3, 2)}-{s.Substring(5, 3)}-{s.Substring(8, 4)}-{s.Substring(12, 1)}"
            End Select
        End If
        Return s
    End Function

    Public Function solicitarLibros(Optional filtro As String = "") As List(Of Libro)
        Return _libroDAL.buscarLibros(filtro)
    End Function

    Public Function solicitarLibrosDisponibles() As List(Of (id As Integer, isbn As Long, titulo As String, autor As String))
        Return _libroDAL.buscarListaLibrosDisponibles()
    End Function

    ''' <summary>
    ''' Obtiene una entidad Libro por su identificador único.
    ''' </summary>
    Public Function buscarLibroPorId(idLibro As Integer) As Libro
        If idLibro <= 0 Then
            Throw New ArgumentException("El identificador del libro debe ser mayor a cero.")
        End If
        Return _libroDAL.buscarLibroPorId(idLibro)
    End Function

    ''' <summary>
    ''' Verifica si un ISBN ya está registrado para otro libro.
    ''' </summary>
    Public Function existeIsbn(isbn As Long, Optional idExcluir As Integer = 0) As Boolean
        Return _libroDAL.existeIsbn(isbn, idExcluir)
    End Function

    ''' <summary>
    ''' Obtiene la cantidad de préstamos activos (no devueltos) asociados al libro.
    ''' </summary>
    Public Function contarPrestamosActivos(idLibro As Integer) As Integer
        If idLibro <= 0 Then Return 0
        Return _libroDAL.contarPrestamosActivos(idLibro)
    End Function

    ''' <summary>
    ''' Valida las reglas de negocio y actualiza los datos y stock de un libro existente en el catálogo.
    ''' El stock disponible se recalcula automáticamente (stockTotal - prestamosActivos) para evitar desfases con la realidad física.
    ''' </summary>
    Public Function modificarLibro(idLibro As Integer,
                                   isbnTexto As String,
                                   titulo As String,
                                   autor As String,
                                   editorial As String,
                                   stockTotalTexto As String,
                                   Optional stockDisponibleTexto As String = "") As Libro
        If idLibro <= 0 Then
            Throw New ArgumentException("Debe seleccionar un libro válido para modificar.")
        End If

        ' 1. Validación de ISBN
        If String.IsNullOrWhiteSpace(isbnTexto) Then
            Throw New ArgumentException("El código ISBN es obligatorio.")
        End If

        Dim isbnLimpio As String = isbnTexto.Replace("-", "").Replace(" ", "").Trim()
        Dim isbnNum As Long
        If isbnLimpio.Length <> 13 OrElse Not Long.TryParse(isbnLimpio, isbnNum) OrElse isbnNum <= 0 Then
            Throw New ArgumentException("El ISBN debe estar compuesto por exactamente 13 dígitos numéricos.")
        End If

        ' 2. Unicidad de ISBN
        If _libroDAL.existeIsbn(isbnNum, idLibro) Then
            Throw New InvalidOperationException($"Ya existe otro libro registrado en el catálogo con el ISBN {FormatearIsbn(isbnNum)}.")
        End If

        ' 3. Validación de Título
        If String.IsNullOrWhiteSpace(titulo) OrElse titulo.Trim().Length < 2 Then
            Throw New ArgumentException("El título del libro es obligatorio y debe tener al menos 2 caracteres.")
        End If

        ' 4. Validación de Autor
        If String.IsNullOrWhiteSpace(autor) OrElse autor.Trim().Length < 2 Then
            Throw New ArgumentException("El nombre del autor es obligatorio y debe tener al menos 2 caracteres.")
        End If

        ' 5. Validación de Stock Total numérico
        Dim stockTotal As Integer
        If Not Integer.TryParse(stockTotalTexto.Trim(), stockTotal) OrElse stockTotal < 0 Then
            Throw New ArgumentException("El stock total debe ser un número entero mayor o igual a 0.")
        End If

        ' 6. Verificación de préstamos activos y cálculo automático del stock disponible
        Dim prestamosActivos As Integer = _libroDAL.contarPrestamosActivos(idLibro)
        If stockTotal < prestamosActivos Then
            Throw New InvalidOperationException($"No se puede reducir el stock total a {stockTotal} porque existen actualmente {prestamosActivos} ejemplar(es) en préstamo activo.")
        End If

        ' El stock disponible se calcula estrictamente para evitar desfases con la realidad
        Dim stockDisponibleCalculado As Integer = stockTotal - prestamosActivos

        ' 7. Persistencia a través de la DAL
        Dim libroExistente As Libro = _libroDAL.buscarLibroPorId(idLibro)
        If libroExistente Is Nothing Then
            Throw New InvalidOperationException("El libro a modificar no fue encontrado en la base de datos.")
        End If

        libroExistente.isbn = isbnNum
        libroExistente.titulo = titulo.Trim()
        libroExistente.autor = autor.Trim()
        libroExistente.editorial = If(String.IsNullOrWhiteSpace(editorial), Nothing, editorial.Trim())
        libroExistente.stockTotal = stockTotal
        libroExistente.stockDisponible = stockDisponibleCalculado

        Dim actualizado As Boolean = _libroDAL.actualizarLibro(libroExistente)
        If Not actualizado Then
            Throw New ApplicationException("No se pudieron guardar los cambios del libro en la base de datos.")
        End If

        Return libroExistente
    End Function

    ''' <summary>
    ''' Valida los datos y registra un nuevo libro en el catálogo a través de la capa DAL.
    ''' El stock disponible se asume automáticamente igual al stock total ingresado.
    ''' </summary>
    Public Function registrarLibro(isbnTexto As String,
                                   titulo As String,
                                   autor As String,
                                   editorial As String,
                                   stockTotalTexto As String,
                                   Optional stockDisponibleTexto As String = "") As Libro
        ' 1. Validación de ISBN
        If String.IsNullOrWhiteSpace(isbnTexto) Then
            Throw New ArgumentException("El código ISBN es obligatorio.")
        End If

        Dim isbnLimpio As String = isbnTexto.Replace("-", "").Replace(" ", "").Trim()
        Dim isbnNum As Long
        If isbnLimpio.Length <> 13 OrElse Not Long.TryParse(isbnLimpio, isbnNum) OrElse isbnNum <= 0 Then
            Throw New ArgumentException("El ISBN debe estar compuesto por exactamente 13 dígitos numéricos.")
        End If

        ' 2. Unicidad de ISBN
        If _libroDAL.existeIsbn(isbnNum) Then
            Throw New InvalidOperationException($"Ya existe un libro registrado en el catálogo con el ISBN {FormatearIsbn(isbnNum)}.")
        End If

        ' 3. Validación de Título
        If String.IsNullOrWhiteSpace(titulo) OrElse titulo.Trim().Length < 2 Then
            Throw New ArgumentException("El título del libro es obligatorio y debe tener al menos 2 caracteres.")
        End If

        ' 4. Validación de Autor
        If String.IsNullOrWhiteSpace(autor) OrElse autor.Trim().Length < 2 Then
            Throw New ArgumentException("El nombre del autor es obligatorio y debe tener al menos 2 caracteres.")
        End If

        ' 5. Validación de Stock Total numérico (el stock disponible inicial siempre es igual al stock total)
        Dim stockTotal As Integer
        If Not Integer.TryParse(stockTotalTexto.Trim(), stockTotal) OrElse stockTotal < 0 Then
            Throw New ArgumentException("El stock total debe ser un número entero mayor o igual a 0.")
        End If

        ' 6. Creación y persistencia (StockDisponible = StockTotal por definición de alta)
        Dim nuevoLibro As New Libro() With {
            .isbn = isbnNum,
            .titulo = titulo.Trim(),
            .autor = autor.Trim(),
            .editorial = If(String.IsNullOrWhiteSpace(editorial), Nothing, editorial.Trim()),
            .stockTotal = stockTotal,
            .stockDisponible = stockTotal
        }

        Dim nuevoId As Integer = _libroDAL.guardarLibro(nuevoLibro)
        nuevoLibro.idLibro = nuevoId

        Return nuevoLibro
    End Function

    ''' <summary>
    ''' Valida las condiciones de negocio y elimina un libro del catálogo.
    ''' Si posee préstamos activos, la eliminación es rechazada para preservar la consistencia.
    ''' </summary>
    Public Function eliminarLibro(idLibro As Integer) As Boolean
        If idLibro <= 0 Then
            Throw New ArgumentException("Debe seleccionar un libro válido para eliminar.")
        End If

        Dim libro As Libro = _libroDAL.buscarLibroPorId(idLibro)
        If libro Is Nothing Then
            Throw New InvalidOperationException("El libro a eliminar no fue encontrado en el catálogo.")
        End If

        Dim prestamosActivos As Integer = _libroDAL.contarPrestamosActivos(idLibro)
        If prestamosActivos > 0 Then
            Throw New InvalidOperationException($"No se puede eliminar el libro '{libro.titulo}' porque posee {prestamosActivos} ejemplar(es) actualmente en préstamo. Registre la devolución de todos los ejemplares antes de eliminarlo.")
        End If

        Dim eliminado As Boolean = _libroDAL.eliminarLibro(idLibro)
        If Not eliminado Then
            Throw New ApplicationException("No se pudo eliminar el libro de la base de datos.")
        End If

        Return True
    End Function

End Class
