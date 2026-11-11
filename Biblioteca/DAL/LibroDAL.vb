Imports Microsoft.Data.Sqlite

Public Class LibroDAL

    Public Function buscarLibros(Optional filtro As String = "") As List(Of Libro)
        Dim lista As New List(Of Libro)()

        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT idLibro, isbn, titulo, autor, editorial, stockTotal, stockDisponible FROM Libros"

            If Not String.IsNullOrWhiteSpace(filtro) Then
                query &= " WHERE (CAST(isbn AS TEXT) LIKE @filtro OR titulo LIKE @filtro OR autor LIKE @filtro)"
            End If

            query &= " ORDER BY idLibro ASC;"

            Using comando As New SqliteCommand(query, conexion)
                If Not String.IsNullOrWhiteSpace(filtro) Then
                    comando.Parameters.AddWithValue("@filtro", "%" & filtro.Trim() & "%")
                End If

                Using reader As SqliteDataReader = comando.ExecuteReader()
                    While reader.Read()
                        Dim l As New Libro() With {
                            .idLibro = reader.GetInt32(0),
                            .isbn = reader.GetInt64(1),
                            .titulo = reader.GetString(2),
                            .autor = reader.GetString(3),
                            .editorial = If(reader.IsDBNull(4), Nothing, reader.GetString(4)),
                            .stockTotal = reader.GetInt32(5),
                            .stockDisponible = reader.GetInt32(6)
                        }
                        lista.Add(l)
                    End While
                End Using
            End Using
        End Using

        Return lista
    End Function

    Public Function buscarListaLibrosDisponibles() As List(Of (id As Integer, isbn As Long, titulo As String, autor As String))
        Dim lista As New List(Of (id As Integer, isbn As Long, titulo As String, autor As String))()

        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT idLibro, isbn, titulo, autor FROM Libros WHERE stockDisponible > 0 ORDER BY titulo ASC;"

            Using comando As New SqliteCommand(query, conexion)
                Using reader As SqliteDataReader = comando.ExecuteReader()
                    While reader.Read()
                        lista.Add((reader.GetInt32(0), reader.GetInt64(1), reader.GetString(2), reader.GetString(3)))
                    End While
                End Using
            End Using
        End Using

        Return lista
    End Function

    ''' <summary>
    ''' Obtiene un libro del catálogo a partir de su clave primaria.
    ''' </summary>
    Public Function buscarLibroPorId(idLibro As Integer) As Libro
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT idLibro, isbn, titulo, autor, editorial, stockTotal, stockDisponible FROM Libros WHERE idLibro = @idLibro LIMIT 1;"

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@idLibro", idLibro)

                Using reader As SqliteDataReader = comando.ExecuteReader()
                    If reader.Read() Then
                        Return MapearLibro(reader)
                    End If
                End Using
            End Using
        End Using

        Return Nothing
    End Function

    ''' <summary>
    ''' Actualiza los datos y niveles de stock de un libro existente en la base de datos.
    ''' </summary>
    Public Function actualizarLibro(libro As Libro) As Boolean
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "
                UPDATE Libros
                SET isbn = @isbn,
                    titulo = @titulo,
                    autor = @autor,
                    editorial = @editorial,
                    stockTotal = @stockTotal,
                    stockDisponible = @stockDisponible
                WHERE idLibro = @idLibro;
            "

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@idLibro", libro.idLibro)
                comando.Parameters.AddWithValue("@isbn", libro.isbn)
                comando.Parameters.AddWithValue("@titulo", libro.titulo.Trim())
                comando.Parameters.AddWithValue("@autor", libro.autor.Trim())
                comando.Parameters.AddWithValue("@editorial", If(String.IsNullOrWhiteSpace(libro.editorial), DBNull.Value, CObj(libro.editorial.Trim())))
                comando.Parameters.AddWithValue("@stockTotal", libro.stockTotal)
                comando.Parameters.AddWithValue("@stockDisponible", libro.stockDisponible)

                Dim filasAfectadas As Integer = comando.ExecuteNonQuery()
                Return filasAfectadas > 0
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Verifica si ya existe otro libro con el mismo ISBN en el catálogo.
    ''' Permite excluir un idLibro específico (útil para edición).
    ''' </summary>
    Public Function existeIsbn(isbn As Long, Optional idExcluir As Integer = 0) As Boolean
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT COUNT(*) FROM Libros WHERE isbn = @isbn AND idLibro <> @idExcluir;"

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@isbn", isbn)
                comando.Parameters.AddWithValue("@idExcluir", idExcluir)

                Dim total As Long = Convert.ToInt64(comando.ExecuteScalar())
                Return total > 0
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Cuenta la cantidad de préstamos activos (no devueltos) asociados a un libro.
    ''' </summary>
    Public Function contarPrestamosActivos(idLibro As Integer) As Integer
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT COUNT(*) FROM Prestamos WHERE idLibro = @idLibro AND fueDevuelto = 0;"

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@idLibro", idLibro)

                Dim total As Long = Convert.ToInt64(comando.ExecuteScalar())
                Return CInt(total)
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Inserta un nuevo libro en la base de datos y retorna su idLibro autogenerado.
    ''' </summary>
    Public Function guardarLibro(libro As Libro) As Integer
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "
                INSERT INTO Libros (isbn, titulo, autor, editorial, stockTotal, stockDisponible)
                VALUES (@isbn, @titulo, @autor, @editorial, @stockTotal, @stockDisponible);
                SELECT last_insert_rowid();
            "

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@isbn", libro.isbn)
                comando.Parameters.AddWithValue("@titulo", libro.titulo.Trim())
                comando.Parameters.AddWithValue("@autor", libro.autor.Trim())
                comando.Parameters.AddWithValue("@editorial", If(String.IsNullOrWhiteSpace(libro.editorial), DBNull.Value, CObj(libro.editorial.Trim())))
                comando.Parameters.AddWithValue("@stockTotal", libro.stockTotal)
                comando.Parameters.AddWithValue("@stockDisponible", libro.stockDisponible)

                Dim nuevoId As Long = Convert.ToInt64(comando.ExecuteScalar())
                libro.idLibro = Convert.ToInt32(nuevoId)
                Return libro.idLibro
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Elimina un libro del catálogo y su historial de préstamos devueltos en una transacción atómica.
    ''' Si posee préstamos activos en curso, no se elimina y retorna False.
    ''' </summary>
    Public Function eliminarLibro(idLibro As Integer) As Boolean
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Using transaccion = conexion.BeginTransaction()
                Try
                    ' 1. Verificar si existen préstamos activos
                    Dim queryCheck As String = "SELECT COUNT(*) FROM Prestamos WHERE idLibro = @idLibro AND fueDevuelto = 0;"
                    Using cmdCheck As New SqliteCommand(queryCheck, conexion, transaccion)
                        cmdCheck.Parameters.AddWithValue("@idLibro", idLibro)
                        Dim activos As Long = Convert.ToInt64(cmdCheck.ExecuteScalar())
                        If activos > 0 Then
                            transaccion.Rollback()
                            Return False
                        End If
                    End Using

                    ' 2. Eliminar préstamos devueltos históricos para mantener la integridad de FK
                    Dim queryDeletePrestamos As String = "DELETE FROM Prestamos WHERE idLibro = @idLibro AND fueDevuelto = 1;"
                    Using cmdPrestamos As New SqliteCommand(queryDeletePrestamos, conexion, transaccion)
                        cmdPrestamos.Parameters.AddWithValue("@idLibro", idLibro)
                        cmdPrestamos.ExecuteNonQuery()
                    End Using

                    ' 3. Eliminar el registro del libro
                    Dim queryDeleteLibro As String = "DELETE FROM Libros WHERE idLibro = @idLibro;"
                    Dim filasAfectadas As Integer = 0
                    Using cmdLibro As New SqliteCommand(queryDeleteLibro, conexion, transaccion)
                        cmdLibro.Parameters.AddWithValue("@idLibro", idLibro)
                        filasAfectadas = cmdLibro.ExecuteNonQuery()
                    End Using

                    transaccion.Commit()
                    Return filasAfectadas > 0

                Catch
                    transaccion.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Function

    Private Shared Function MapearLibro(reader As SqliteDataReader) As Libro
        Return New Libro() With {
            .idLibro = reader.GetInt32(0),
            .isbn = reader.GetInt64(1),
            .titulo = reader.GetString(2),
            .autor = reader.GetString(3),
            .editorial = If(reader.IsDBNull(4), Nothing, reader.GetString(4)),
            .stockTotal = reader.GetInt32(5),
            .stockDisponible = reader.GetInt32(6)
        }
    End Function

End Class

