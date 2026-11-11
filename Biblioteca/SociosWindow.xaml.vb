Imports System.Text.RegularExpressions
Imports System.Threading

Class SociosWindow

    Private ReadOnly _socioBLL As New SocioBLL()
    Private _cancellationTokenSource As CancellationTokenSource

    Private Async Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        Await CargarListaSociosAsync()
    End Sub

    ' =========================== NAVEGACIÓN Y PESTAÑAS ===========================

    ''' <summary>
    ''' Muestra la pantalla principal con la lista de socios y su estado.
    ''' </summary>
    Public Async Sub MostrarModuloListaSocios()
        ActualizarNavegacionLateral(esListaSocios:=True)

        gridRegistrarSocio.Visibility = Visibility.Collapsed
        gridListaSocios.Visibility = Visibility.Visible
        scrollPrincipal.ScrollToTop()

        Await CargarListaSociosAsync(txtBuscarSocios.Text.Trim())
    End Sub

    ''' <summary>
    ''' Muestra el formulario para registrar un nuevo socio.
    ''' </summary>
    Public Sub MostrarModuloRegistrarSocio()
        ActualizarNavegacionLateral(esListaSocios:=False)

        gridListaSocios.Visibility = Visibility.Collapsed
        gridRegistrarSocio.Visibility = Visibility.Visible
        scrollPrincipal.ScrollToTop()

        txtNuevoNombre.Focus()
    End Sub

    Private Sub BtnListaSocios_Click(sender As Object, e As RoutedEventArgs)
        MostrarModuloListaSocios()
    End Sub

    Private Sub BtnNavRegistrarSocio_Click(sender As Object, e As RoutedEventArgs)
        MostrarModuloRegistrarSocio()
    End Sub

    Private Sub BtnVolverListaSocios_Click(sender As Object, e As RoutedEventArgs)
        MostrarModuloListaSocios()
    End Sub

    Private Sub ActualizarNavegacionLateral(esListaSocios As Boolean)
        Dim verdeActivo As Brush = New SolidColorBrush(Color.FromRgb(&H1D, &H3B, &H31))
        Dim verdeFondoPill As Brush = New SolidColorBrush(Color.FromRgb(&HEB, &HF3, &HEF))
        Dim grisTexto As Brush = New SolidColorBrush(Color.FromRgb(&H4B, &H55, &H63))

        If esListaSocios Then
            btnNavListaSocios.Background = verdeFondoPill
            textNavListaSocios.Foreground = verdeActivo
            textNavListaSocios.FontWeight = FontWeights.SemiBold
            iconNavListaSocios.Stroke = verdeActivo

            btnNavRegistrarSocio.Background = Brushes.Transparent
            textNavRegistrarSocio.Foreground = grisTexto
            textNavRegistrarSocio.FontWeight = FontWeights.Normal
            iconNavRegistrarSocio.Stroke = grisTexto
        Else
            btnNavRegistrarSocio.Background = verdeFondoPill
            textNavRegistrarSocio.Foreground = verdeActivo
            textNavRegistrarSocio.FontWeight = FontWeights.SemiBold
            iconNavRegistrarSocio.Stroke = verdeActivo

            btnNavListaSocios.Background = Brushes.Transparent
            textNavListaSocios.Foreground = grisTexto
            textNavListaSocios.FontWeight = FontWeights.Normal
            iconNavListaSocios.Stroke = grisTexto
        End If
    End Sub

    ' =========================== BÚSQUEDA Y CARGA ASÍNCRONA ===========================

    Private Async Sub TxtBuscarSocios_TextChanged(sender As Object, e As TextChangedEventArgs)
        txtPlaceholderBuscarSocios.Visibility = If(String.IsNullOrEmpty(txtBuscarSocios.Text), Visibility.Visible, Visibility.Collapsed)

        Dim textoBusqueda As String = txtBuscarSocios.Text.Trim()

        If _cancellationTokenSource IsNot Nothing Then
            _cancellationTokenSource.Cancel()
            _cancellationTokenSource.Dispose()
        End If

        _cancellationTokenSource = New CancellationTokenSource()
        Dim token As CancellationToken = _cancellationTokenSource.Token

        Try
            Await Task.Delay(400, token)
            Await CargarListaSociosAsync(textoBusqueda)
        Catch ex As TaskCanceledException
            ' Cancelado porque el usuario continuó tecleando
        Catch ex As Exception
            MessageBox.Show("Error al buscar socios: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Carga los socios desde la capa BLL de forma asíncrona y refresca los elementos visuales.
    ''' </summary>
    Private Async Function CargarListaSociosAsync(Optional filtro As String = "") As Task
        Try
            Dim lista = Await Task.Run(Function() _socioBLL.solicitarTodosLosSocios(filtro))
            icSocios.ItemsSource = lista
            lblContadorSocios.Text = $"{lista.Count} socios registrados"

            If lista.Count = 0 Then
                panelSinResultados.Visibility = Visibility.Visible
            Else
                panelSinResultados.Visibility = Visibility.Collapsed
            End If
        Catch ex As Exception
            MessageBox.Show("Error al cargar la lista de socios: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Function

    ' =========================== ACCIONES DE ESTADO (ACTIVAR / DESACTIVAR) ===========================

    Private Async Sub BtnDesactivarSocio_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = TryCast(sender, Button)
        Dim socioItem As SocioItem = TryCast(btn?.Tag, SocioItem)

        If socioItem Is Nothing Then Return

        Dim confirmacion As MessageBoxResult = MessageBox.Show(
            $"¿Estás seguro de que deseas desactivar al socio '{socioItem.NombreCompleto}'?{Environment.NewLine}{Environment.NewLine}El socio figurará como Inhabilitado y no podrá solicitar nuevos préstamos.",
            "Confirmar Desactivación",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question)

        If confirmacion = MessageBoxResult.Yes Then
            Try
                Dim exito As Boolean = Await Task.Run(Function()
                    Return _socioBLL.cambiarEstadoSocio(socioItem.IdSocio, EstadoSocio.Inactivo)
                End Function)

                If exito Then
                    Await CargarListaSociosAsync(txtBuscarSocios.Text.Trim())
                Else
                    MessageBox.Show("No se pudo desactivar el socio.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
                End If
            Catch ex As Exception
                MessageBox.Show("Error al desactivar el socio: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Async Sub BtnActivarSocio_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = TryCast(sender, Button)
        Dim socioItem As SocioItem = TryCast(btn?.Tag, SocioItem)

        If socioItem Is Nothing Then Return

        Dim confirmacion As MessageBoxResult = MessageBox.Show(
            $"¿Deseas reactivar al socio '{socioItem.NombreCompleto}'?{Environment.NewLine}{Environment.NewLine}El socio figurará como Habilitado y podrá volver a realizar préstamos.",
            "Confirmar Activación",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question)

        If confirmacion = MessageBoxResult.Yes Then
            Try
                Dim exito As Boolean = Await Task.Run(Function()
                    Return _socioBLL.cambiarEstadoSocio(socioItem.IdSocio, EstadoSocio.Activo)
                End Function)

                If exito Then
                    Await CargarListaSociosAsync(txtBuscarSocios.Text.Trim())
                Else
                    MessageBox.Show("No se pudo activar el socio.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
                End If
            Catch ex As Exception
                MessageBox.Show("Error al activar el socio: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    ' =========================== FORMULARIO ALTA DE SOCIO ===========================

    Private Sub TxtNuevoDni_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        ' Solo permitir números
        Dim regex As New Regex("[^0-9]+")
        e.Handled = regex.IsMatch(e.Text)
    End Sub

    Private Async Sub BtnGuardarSocio_Click(sender As Object, e As RoutedEventArgs)
        Dim nombre As String = txtNuevoNombre.Text.Trim()
        Dim dniTexto As String = txtNuevoDni.Text.Trim()
        Dim email As String = txtNuevoEmail.Text.Trim()
        Dim telefono As String = txtNuevoTelefono.Text.Trim()

        If String.IsNullOrWhiteSpace(nombre) Then
            MessageBox.Show("Por favor ingresa el nombre y apellido del socio.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning)
            txtNuevoNombre.Focus()
            Return
        End If

        Dim dni As Long
        If Not Long.TryParse(dniTexto, dni) OrElse dni <= 0 Then
            MessageBox.Show("Por favor ingresa un número de DNI válido.", "Validación", MessageBoxButton.OK, MessageBoxImage.Warning)
            txtNuevoDni.Focus()
            Return
        End If

        Try
            Dim nuevoSocio As Socio = Await Task.Run(Function()
                Return _socioBLL.registrarSocio(nombre, dni, email, telefono)
            End Function)

            MessageBox.Show(
                $"Socio '{nuevoSocio.nombreCompleto}' registrado exitosamente con estado Habilitado.",
                "Registro Exitoso",
                MessageBoxButton.OK,
                MessageBoxImage.Information)

            ' Limpiar formulario
            txtNuevoNombre.Clear()
            txtNuevoDni.Clear()
            txtNuevoEmail.Clear()
            txtNuevoTelefono.Clear()

            ' Volver al listado
            MostrarModuloListaSocios()

        Catch ex As InvalidOperationException
            MessageBox.Show(ex.Message, "Aviso de Validación", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error al guardar el socio: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' =========================== NAVEGACIÓN Y CIERRE ===========================

    Private Sub BtnVolverMenu_Click(sender As Object, e As RoutedEventArgs)
        CerrarYVolverAlMenu()
    End Sub

    Private Sub SociosWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If Me.Owner IsNot Nothing AndAlso Not Me.Owner.IsVisible Then
            Me.Owner.Show()
        End If
    End Sub

    Private Sub CerrarYVolverAlMenu()
        If Me.Owner IsNot Nothing Then
            Me.Owner.Show()
        End If
        Me.Close()
    End Sub

End Class

