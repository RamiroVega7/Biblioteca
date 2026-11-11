Imports Microsoft.Data.Sqlite

Public Class SocioDAL

    ' =========================================================================
    ' MÉTODOS PÚBLICOS DE CONSULTA (CLEAN CODE)
    ' =========================================================================

    ''' <summary>
    ''' Obtiene exclusivamente los socios activos/habilitados (para préstamos y operaciones regulares).
    ''' </summary>
    Public Function buscarSociosHabilitados(Optional filtro As String = "") As List(Of Socio)
        Dim query As String = "SELECT idSocio, nombreCompleto, dni, email, telefono, estado FROM Socios WHERE estado = 0"

        If Not String.IsNullOrWhiteSpace(filtro) Then
            query &= " AND (CAST(dni AS TEXT) LIKE @filtro OR LOWER(nombreCompleto) LIKE @filtro)"
        End If

        query &= " ORDER BY nombreCompleto ASC;"

        Return EjecutarConsultaSocios(query, filtro)
    End Function

    ''' <summary>
    ''' Obtiene todos los socios (activos e inactivos) para la pantalla de Administración y auditoría.
    ''' </summary>
    Public Function buscarTodosLosSocios(Optional filtro As String = "") As List(Of Socio)
        Dim query As String = "SELECT idSocio, nombreCompleto, dni, email, telefono, estado FROM Socios WHERE 1 = 1"

        If Not String.IsNullOrWhiteSpace(filtro) Then
            query &= " AND (CAST(dni AS TEXT) LIKE @filtro OR LOWER(nombreCompleto) LIKE @filtro OR LOWER(COALESCE(email, '')) LIKE @filtro)"
        End If

        query &= " ORDER BY idSocio ASC;"

        Return EjecutarConsultaSocios(query, filtro)
    End Function

    ''' <summary>
    ''' Proyección ligera y optimizada en SQLite: solo extrae las columnas requeridas (id, nombre, dni) de socios activos.
    ''' </summary>
    Public Function buscarTuplasSociosHabilitados(Optional filtro As String = "") As List(Of (id As Integer, nombreCompleto As String, dni As Long))
        Dim lista As New List(Of (id As Integer, nombreCompleto As String, dni As Long))()

        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT idSocio, nombreCompleto, dni FROM Socios WHERE estado = 0"

            If Not String.IsNullOrWhiteSpace(filtro) Then
                query &= " AND (CAST(dni AS TEXT) LIKE @filtro OR LOWER(nombreCompleto) LIKE @filtro)"
            End If

            query &= " ORDER BY nombreCompleto ASC;"

            Using comando As New SqliteCommand(query, conexion)
                If Not String.IsNullOrWhiteSpace(filtro) Then
                    comando.Parameters.AddWithValue("@filtro", "%" & filtro.Trim().ToLowerInvariant() & "%")
                End If

                Using reader As SqliteDataReader = comando.ExecuteReader()
                    While reader.Read()
                        lista.Add((reader.GetInt32(0), reader.GetString(1), reader.GetInt64(2)))
                    End While
                End Using
            End Using
        End Using

        Return lista
    End Function

    ''' <summary>
    ''' Busca un socio por su identificador primario único.
    ''' </summary>
    Public Function buscarSocioPorId(idSocio As Integer) As Socio
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT idSocio, nombreCompleto, dni, email, telefono, estado FROM Socios WHERE idSocio = @idSocio LIMIT 1;"

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@idSocio", idSocio)

                Using reader As SqliteDataReader = comando.ExecuteReader()
                    If reader.Read() Then
                        Return MapearSocio(reader)
                    End If
                End Using
            End Using
        End Using

        Return Nothing
    End Function

    ' =========================================================================
    ' MÉTODOS DE ESCRITURA Y VALIDACIÓN
    ' =========================================================================

    ''' <summary>
    ''' Actualiza el estado de un socio (Activo / Inactivo).
    ''' </summary>
    Public Function cambiarEstadoSocio(idSocio As Integer, nuevoEstado As EstadoSocio) As Boolean
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "UPDATE Socios SET estado = @estado WHERE idSocio = @idSocio;"

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@estado", CInt(nuevoEstado))
                comando.Parameters.AddWithValue("@idSocio", idSocio)

                Dim filasAfectadas As Integer = comando.ExecuteNonQuery()
                Return filasAfectadas > 0
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Inserta un nuevo socio en la base de datos y retorna el ID generado.
    ''' </summary>
    Public Function guardarSocio(socio As Socio) As Integer
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "
                INSERT INTO Socios (nombreCompleto, dni, email, telefono, estado)
                VALUES (@nombreCompleto, @dni, @email, @telefono, @estado);
                SELECT last_insert_rowid();"

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@nombreCompleto", socio.nombreCompleto.Trim())
                comando.Parameters.AddWithValue("@dni", socio.dni)
                comando.Parameters.AddWithValue("@email", If(String.IsNullOrWhiteSpace(socio.email), DBNull.Value, socio.email.Trim()))
                comando.Parameters.AddWithValue("@telefono", If(String.IsNullOrWhiteSpace(socio.telefono), DBNull.Value, socio.telefono.Trim()))
                comando.Parameters.AddWithValue("@estado", CInt(socio.estado))

                Dim nuevoId As Long = Convert.ToInt64(comando.ExecuteScalar())
                socio.idSocio = CInt(nuevoId)
                Return socio.idSocio
            End Using
        End Using
    End Function

    ''' <summary>
    ''' Verifica si ya existe un socio con el DNI especificado.
    ''' </summary>
    Public Function existeDni(dni As Long, Optional idExcluir As Integer = 0) As Boolean
        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Dim query As String = "SELECT COUNT(*) FROM Socios WHERE dni = @dni AND idSocio <> @idExcluir;"

            Using comando As New SqliteCommand(query, conexion)
                comando.Parameters.AddWithValue("@dni", dni)
                comando.Parameters.AddWithValue("@idExcluir", idExcluir)

                Dim total As Long = Convert.ToInt64(comando.ExecuteScalar())
                Return total > 0
            End Using
        End Using
    End Function

    ' =========================================================================
    ' MÉTODOS PRIVADOS AUXILIARES (DRY)
    ' =========================================================================

    ''' <summary>
    ''' Centraliza la ejecución de consultas que retornan entidades Socio completas.
    ''' </summary>
    Private Function EjecutarConsultaSocios(query As String, filtro As String) As List(Of Socio)
        Dim lista As New List(Of Socio)()

        Using conexion As New SqliteConnection(DatabaseHelper.ConnectionString)
            conexion.Open()

            Using comando As New SqliteCommand(query, conexion)
                If Not String.IsNullOrWhiteSpace(filtro) Then
                    comando.Parameters.AddWithValue("@filtro", "%" & filtro.Trim().ToLowerInvariant() & "%")
                End If

                Using reader As SqliteDataReader = comando.ExecuteReader()
                    While reader.Read()
                        lista.Add(MapearSocio(reader))
                    End While
                End Using
            End Using
        End Using

        Return lista
    End Function

    ''' <summary>
    ''' Mapea de forma segura y centralizada un registro de SqliteDataReader a una entidad Socio.
    ''' </summary>
    Private Shared Function MapearSocio(reader As SqliteDataReader) As Socio
        Return New Socio() With {
            .idSocio = reader.GetInt32(0),
            .nombreCompleto = reader.GetString(1),
            .dni = reader.GetInt64(2),
            .email = If(reader.IsDBNull(3), Nothing, reader.GetString(3)),
            .telefono = If(reader.IsDBNull(4), Nothing, reader.GetString(4)),
            .estado = CType(reader.GetInt32(5), EstadoSocio)
        }
    End Function

End Class
