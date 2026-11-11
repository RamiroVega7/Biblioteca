Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class SocioBLL

    Private ReadOnly _socioDAL As New SocioDAL()
    Private Shared ReadOnly _culturaEs As New CultureInfo("es-AR")

    ''' <summary>
    ''' Solicita la lista de socios habilitados/activos (para préstamos y operaciones regulares).
    ''' </summary>
    Public Function solicitarSociosHabilitados(Optional filtro As String = "") As List(Of SocioItem)
        Dim socios As List(Of Socio) = _socioDAL.buscarSociosHabilitados(filtro)
        Return MapearListaSociosItem(socios)
    End Function

    ''' <summary>
    ''' Solicita la lista completa de socios (activos e inhabilitados) para administración y auditoría.
    ''' </summary>
    Public Function solicitarTodosLosSocios(Optional filtro As String = "") As List(Of SocioItem)
        Dim socios As List(Of Socio) = _socioDAL.buscarTodosLosSocios(filtro)
        Return MapearListaSociosItem(socios)
    End Function

    ''' <summary>
    ''' Proyección ligera de socios activos para ComboBox y controles de selección rápida.
    ''' </summary>
    Public Function solicitarSociosParaCombo(Optional filtro As String = "") As List(Of (id As Integer, nombreCompleto As String, dni As Long))
        Return _socioDAL.buscarTuplasSociosHabilitados(filtro)
    End Function

    ' =========================================================================
    ' Mapeo común (DRY - Don't Repeat Yourself)
    ' =========================================================================

    ''' <summary>
    ''' Transforma una lista de entidades Socio en una lista de modelos de vista SocioItem.
    ''' </summary>
    Private Shared Function MapearListaSociosItem(socios As IEnumerable(Of Socio)) As List(Of SocioItem)
        Dim lista As New List(Of SocioItem)()
        For Each s In socios
            lista.Add(MapearASocioItem(s))
        Next
        Return lista
    End Function

    ''' <summary>
    ''' Transforma una entidad de dominio Socio a un modelo de vista SocioItem con formato regional.
    ''' </summary>
    Private Shared Function MapearASocioItem(s As Socio) As SocioItem
        Return New SocioItem() With {
            .IdSocio = s.idSocio,
            .NombreCompleto = s.nombreCompleto,
            .Dni = s.dni,
            .DniFormateado = s.dni.ToString("N0", _culturaEs),
            .Email = If(String.IsNullOrWhiteSpace(s.email), "-", s.email),
            .Telefono = If(String.IsNullOrWhiteSpace(s.telefono), "-", s.telefono),
            .Estado = s.estado
        }
    End Function

    ''' <summary>
    ''' Cambia el estado (Habilitado / Inhabilitado) de un socio.
    ''' </summary>
    Public Function cambiarEstadoSocio(idSocio As Integer, nuevoEstado As EstadoSocio) As Boolean
        If idSocio <= 0 Then
            Throw New ArgumentException("El identificador del socio no es válido.", NameOf(idSocio))
        End If

        Return _socioDAL.cambiarEstadoSocio(idSocio, nuevoEstado)
    End Function

    ''' <summary>
    ''' Valida los datos y registra un nuevo socio activo en el sistema.
    ''' </summary>
    Public Function registrarSocio(nombreCompleto As String, dni As Long, email As String, telefono As String) As Socio
        ' Validaciones de negocio
        If String.IsNullOrWhiteSpace(nombreCompleto) OrElse nombreCompleto.Trim().Length < 3 Then
            Throw New InvalidOperationException("El nombre completo debe tener al menos 3 caracteres.")
        End If

        If dni <= 0 OrElse dni > 999999999 Then
            Throw New InvalidOperationException("El número de DNI ingresado no es válido.")
        End If

        If _socioDAL.existeDni(dni) Then
            Throw New InvalidOperationException($"Ya existe un socio registrado con el DNI {dni.ToString("N0", _culturaEs)}.")
        End If

        If Not String.IsNullOrWhiteSpace(email) Then
            Dim emailRegex As New Regex("^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.IgnoreCase)
            If Not emailRegex.IsMatch(email.Trim()) Then
                Throw New InvalidOperationException("El formato del correo electrónico ingresado no es válido.")
            End If
        End If

        Dim nuevoSocio As New Socio() With {
            .nombreCompleto = nombreCompleto.Trim(),
            .dni = dni,
            .email = If(String.IsNullOrWhiteSpace(email), Nothing, email.Trim()),
            .telefono = If(String.IsNullOrWhiteSpace(telefono), Nothing, telefono.Trim()),
            .estado = EstadoSocio.Activo
        }

        Dim idGenerado As Integer = _socioDAL.guardarSocio(nuevoSocio)
        nuevoSocio.idSocio = idGenerado

        Return nuevoSocio
    End Function

End Class
