Imports Microsoft.Data.Sqlite

Public Class PrestamoDAL

    Public Function guardarPrestamo(prestamo As Prestamo) As Integer
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Using transaccion = conexion.BeginTransaction()
                Try
                    Dim query As String = "
                        INSERT INTO Prestamos (idSocio, idLibro, fechaPrestamo, fechaLimite, fechaDevolucion, fueDevuelto)
                        VALUES (@idSocio, @idLibro, @fechaPrestamo, @fechaLimite, @fechaDevolucion, @fueDevuelto);
                        SELECT last_insert_rowid();
                    "

                    Using comando As New SqliteCommand(query, conexion, transaccion)
                        comando.Parameters.AddWithValue("@idSocio", prestamo.socio.idSocio)
                        comando.Parameters.AddWithValue("@idLibro", prestamo.libro.idLibro)
                        comando.Parameters.AddWithValue("@fechaPrestamo", prestamo.fechaPrestamo.ToString("yyyy-MM-dd"))
                        comando.Parameters.AddWithValue("@fechaLimite", prestamo.fechaLimite.ToString("yyyy-MM-dd"))
                        comando.Parameters.AddWithValue("@fechaDevolucion", If(prestamo.fechaDevolucion.HasValue, CObj(prestamo.fechaDevolucion.Value.ToString("yyyy-MM-dd")), DBNull.Value))
                        comando.Parameters.AddWithValue("@fueDevuelto", If(prestamo.fueDevuelto, 1, 0))

                        Dim nuevoId As Long = Convert.ToInt64(comando.ExecuteScalar())
                        prestamo.idPrestamo = Convert.ToInt32(nuevoId)
                    End Using

                    ' Disminuir el stock disponible del libro prestado
                    Dim queryUpdateStock As String = "UPDATE Libros SET stockDisponible = stockDisponible - 1 WHERE idLibro = @idLibro AND stockDisponible > 0;"
                    Using cmdStock As New SqliteCommand(queryUpdateStock, conexion, transaccion)
                        cmdStock.Parameters.AddWithValue("@idLibro", prestamo.libro.idLibro)
                        cmdStock.ExecuteNonQuery()
                    End Using

                    transaccion.Commit()
                    Return prestamo.idPrestamo

                Catch
                    transaccion.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Function

    Public Function registrarDevolucion(idPrestamo As Integer) As Boolean
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Using transaccion = conexion.BeginTransaction()
                Try
                    Dim idLibro As Integer = 0
                    Dim yaDevuelto As Boolean = False
                    Dim queryCheck As String = "SELECT idLibro, fueDevuelto FROM Prestamos WHERE idPrestamo = @idPrestamo;"

                    Using cmdCheck As New SqliteCommand(queryCheck, conexion, transaccion)
                        cmdCheck.Parameters.AddWithValue("@idPrestamo", idPrestamo)
                        Using reader = cmdCheck.ExecuteReader()
                            If reader.Read() Then
                                idLibro = reader.GetInt32(0)
                                yaDevuelto = (reader.GetInt32(1) = 1)
                            Else
                                Return False
                            End If
                        End Using
                    End Using

                    If yaDevuelto Then
                        Return True
                    End If

                    Dim queryUpdate As String = "
                        UPDATE Prestamos 
                        SET fueDevuelto = 1, fechaDevolucion = @fechaDevolucion 
                        WHERE idPrestamo = @idPrestamo;
                    "
                    Using cmdUpdate As New SqliteCommand(queryUpdate, conexion, transaccion)
                        cmdUpdate.Parameters.AddWithValue("@idPrestamo", idPrestamo)
                        cmdUpdate.Parameters.AddWithValue("@fechaDevolucion", Date.Today.ToString("yyyy-MM-dd"))
                        cmdUpdate.ExecuteNonQuery()
                    End Using

                    ' Restaurar stock disponible del libro
                    Dim queryStock As String = "UPDATE Libros SET stockDisponible = stockDisponible + 1 WHERE idLibro = @idLibro;"
                    Using cmdStock As New SqliteCommand(queryStock, conexion, transaccion)
                        cmdStock.Parameters.AddWithValue("@idLibro", idLibro)
                        cmdStock.ExecuteNonQuery()
                    End Using

                    transaccion.Commit()
                    Return True

                Catch
                    transaccion.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Function

    Public Function buscarPrestamos(Optional filtro As String = "") As List(Of Prestamo)
        Dim lista As New List(Of Prestamo)()

        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "
                SELECT 
                    p.idPrestamo, 
                    p.fechaPrestamo, 
                    p.fechaLimite, 
                    p.fechaDevolucion, 
                    p.fueDevuelto,
                    s.idSocio, 
                    s.nombreCompleto, 
                    s.dni, 
                    s.email, 
                    s.telefono, 
                    s.estado,
                    l.idLibro, 
                    l.isbn, 
                    l.titulo, 
                    l.autor, 
                    l.editorial, 
                    l.stockTotal, 
                    l.stockDisponible
                FROM Prestamos p
                INNER JOIN Socios s ON p.idSocio = s.idSocio
                INNER JOIN Libros l ON p.idLibro = l.idLibro
            "

            If Not String.IsNullOrWhiteSpace(filtro) Then
                query &= " WHERE s.nombreCompleto LIKE @filtro OR l.titulo LIKE @filtro OR l.autor LIKE @filtro"
            End If

            query &= " ORDER BY p.idPrestamo ASC;"

            Using comando As New SqliteCommand(query, conexion)
                If Not String.IsNullOrWhiteSpace(filtro) Then
                    comando.Parameters.AddWithValue("@filtro", "%" & filtro.Trim() & "%")
                End If

                Using reader As SqliteDataReader = comando.ExecuteReader()
                    While reader.Read()
                        Dim socio As New Socio() With {
                            .idSocio = reader.GetInt32(5),
                            .nombreCompleto = reader.GetString(6),
                            .dni = reader.GetInt64(7),
                            .email = If(reader.IsDBNull(8), Nothing, reader.GetString(8)),
                            .telefono = If(reader.IsDBNull(9), Nothing, reader.GetString(9)),
                            .estado = CType(reader.GetInt32(10), EstadoSocio)
                        }

                        Dim libro As New Libro() With {
                            .idLibro = reader.GetInt32(11),
                            .isbn = reader.GetInt64(12),
                            .titulo = reader.GetString(13),
                            .autor = reader.GetString(14),
                            .editorial = If(reader.IsDBNull(15), Nothing, reader.GetString(15)),
                            .stockTotal = reader.GetInt32(16),
                            .stockDisponible = reader.GetInt32(17)
                        }

                        Dim prestamo As New Prestamo() With {
                            .idPrestamo = reader.GetInt32(0),
                            .fechaPrestamo = Date.Parse(reader.GetString(1)),
                            .fechaLimite = Date.Parse(reader.GetString(2)),
                            .fechaDevolucion = If(reader.IsDBNull(3), CType(Nothing, Date?), Date.Parse(reader.GetString(3))),
                            .fueDevuelto = (reader.GetInt32(4) = 1),
                            .socio = socio,
                            .libro = libro
                        }

                        lista.Add(prestamo)
                    End While
                End Using
            End Using
        End Using

        Return lista
    End Function

End Class
