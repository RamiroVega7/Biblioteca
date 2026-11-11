Imports System.Threading

Class PrestamosWindow

    Private ReadOnly _socioBLL As New SocioBLL()
    Private ReadOnly _libroBLL As New LibroBLL()
    Private ReadOnly _prestamoBLL As New PrestamoBLL()

    Private _cancellationTokenSource As CancellationTokenSource
    Private _cancellationTokenSourceLibros As CancellationTokenSource
    Private _cancellationTokenSourcePrestamos As CancellationTokenSource

    Private _socioSeleccionado As SocioItem
    Private _libroSeleccionado As LibroItem

    Private Async Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        Await CargarSociosAsync()
    End Sub

    ' =========================== MÓDULO: LISTA DE PRESTAMOS ===========================

    ''' <summary>
    ''' Abre la pantalla con la lista de préstamos registrados.
    ''' </summary>
    Public Async Sub MostrarModuloListaPrestamos()
        ActualizarNavegacionLateral(esListaPrestamos:=True)

        gridPaso1Socio.Visibility = Visibility.Collapsed
        gridPaso2Libro.Visibility = Visibility.Collapsed
        gridPaso3Plazo.Visibility = Visibility.Collapsed
        gridPasoExito.Visibility = Visibility.Collapsed
        gridListaPrestamos.Visibility = Visibility.Visible
        scrollPrincipal.ScrollToTop()

        Await CargarListaPrestamosAsync()
    End Sub

    Private Async Sub TxtBuscarPrestamos_TextChanged(sender As Object, e As TextChangedEventArgs)
        txtPlaceholderBuscarPrestamos.Visibility = If(String.IsNullOrEmpty(txtBuscarPrestamos.Text), Visibility.Visible, Visibility.Collapsed)

        Dim textoBusqueda As String = txtBuscarPrestamos.Text.Trim()

        If _cancellationTokenSourcePrestamos IsNot Nothing Then
            _cancellationTokenSourcePrestamos.Cancel()
            _cancellationTokenSourcePrestamos.Dispose()
        End If

        _cancellationTokenSourcePrestamos = New CancellationTokenSource()
        Dim token As CancellationToken = _cancellationTokenSourcePrestamos.Token

        Try
            Await Task.Delay(500, token)
            Await CargarListaPrestamosAsync(textoBusqueda)
        Catch ex As TaskCanceledException
            ' Ignorado al seguir escribiendo
        Catch ex As Exception
            MessageBox.Show("Error al buscar préstamos: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Carga los préstamos desde la BLL en segundo plano y actualiza la lista y el contador.
    ''' </summary>
    Private Async Function CargarListaPrestamosAsync(Optional filtro As String = "") As Task
        Try
            Dim lista = Await Task.Run(Function() _prestamoBLL.solicitarListaPrestamos(filtro))
            icListaPrestamos.ItemsSource = lista
            lblContadorListaPrestamos.Text = $"{lista.Count} préstamos registrados"
        Catch ex As Exception
            MessageBox.Show("Error al cargar la lista de préstamos: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Function

    Private Sub BtnListaPrestamos_Click(sender As Object, e As RoutedEventArgs)
        MostrarModuloListaPrestamos()
    End Sub

    Private Async Sub BtnDevolverPrestamo_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = TryCast(sender, Button)
        Dim prestamoItem As PrestamoItem = TryCast(btn?.Tag, PrestamoItem)

        If prestamoItem Is Nothing Then Return

        Dim confirmacion As MessageBoxResult = MessageBox.Show(
            $"¿Deseas registrar la devolución del libro '{prestamoItem.TituloLibro}' prestado a {prestamoItem.NombreSocio}?",
            "Confirmar Devolución",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question)

        If confirmacion = MessageBoxResult.Yes Then
            Try
                Dim exito As Boolean = Await Task.Run(Function()
                    Return _prestamoBLL.registrarDevolucion(prestamoItem.IdPrestamo)
                End Function)

                If exito Then
                    MessageBox.Show(
                        $"Devolución registrada exitosamente.{Environment.NewLine}El ejemplar de '{prestamoItem.TituloLibro}' ha regresado al stock disponible.",
                        "Devolución Exitosa",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information)

                    ' Recargar préstamos y catálogo de libros
                    Await CargarListaPrestamosAsync(txtBuscarPrestamos.Text.Trim())
                    Await CargarLibrosAsync()
                Else
                    MessageBox.Show("No se pudo procesar la devolución.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
                End If

            Catch ex As Exception
                MessageBox.Show("Error al procesar la devolución: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    ' =========================== PASO 1: SELECCIONAR SOCIO ===========================

    Public Sub MostrarPaso1Socios()
        ActualizarNavegacionLateral(esListaPrestamos:=False)

        gridListaPrestamos.Visibility = Visibility.Collapsed
        gridPaso2Libro.Visibility = Visibility.Collapsed
        gridPaso3Plazo.Visibility = Visibility.Collapsed
        gridPasoExito.Visibility = Visibility.Collapsed
        gridPaso1Socio.Visibility = Visibility.Visible
        scrollPrincipal.ScrollToTop()
    End Sub

    Private Sub BtnNavRegistrarPrestamo_Click(sender As Object, e As RoutedEventArgs)
        MostrarPaso1Socios()
    End Sub

    Private Async Sub TxtBuscarDni_TextChanged(sender As Object, e As TextChangedEventArgs)
        txtPlaceholderDni.Visibility = If(String.IsNullOrEmpty(txtBuscarDni.Text), Visibility.Visible, Visibility.Collapsed)

        Dim textoBusqueda As String = txtBuscarDni.Text.Trim()

        If _cancellationTokenSource IsNot Nothing Then
            _cancellationTokenSource.Cancel()
            _cancellationTokenSource.Dispose()
        End If

        _cancellationTokenSource = New CancellationTokenSource()
        Dim token As CancellationToken = _cancellationTokenSource.Token

        Try
            Await Task.Delay(500, token)
            Await CargarSociosAsync(textoBusqueda)
        Catch ex As TaskCanceledException
            ' Ignorado al seguir escribiendo
        Catch ex As Exception
            MessageBox.Show("Error al buscar socio: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Async Function CargarSociosAsync(Optional filtro As String = "") As Task
        Try
            Dim socios = Await Task.Run(Function() _socioBLL.solicitarSociosHabilitados(filtro))
            icSocios.ItemsSource = socios
            lblContadorSocios.Text = $"{socios.Count} socio(s) registrado(s)"
        Catch ex As Exception
            MessageBox.Show("Error al cargar socios: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Function

    Private Sub BtnSeleccionarSocio_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = TryCast(sender, Button)
        Dim socio As SocioItem = TryCast(btn?.Tag, SocioItem)

        If socio IsNot Nothing Then
            _socioSeleccionado = socio
            MostrarPaso2Libros()
        End If
    End Sub

    Private Sub BtnRegistrarNuevoSocio_Click(sender As Object, e As RoutedEventArgs)
        MessageBox.Show("Módulo para registrar un nuevo socio.", "Registrar Socio", MessageBoxButton.OK, MessageBoxImage.Information)
    End Sub

    ' =========================== PASO 2: SELECCIONAR LIBRO ===========================

    Private Async Sub MostrarPaso2Libros()
        ActualizarNavegacionLateral(esListaPrestamos:=False)

        gridListaPrestamos.Visibility = Visibility.Collapsed
        gridPaso1Socio.Visibility = Visibility.Collapsed
        gridPaso3Plazo.Visibility = Visibility.Collapsed
        gridPasoExito.Visibility = Visibility.Collapsed
        gridPaso2Libro.Visibility = Visibility.Visible
        scrollPrincipal.ScrollToTop()
        Await CargarLibrosAsync()
    End Sub

    Private Async Sub TxtBuscarLibro_TextChanged(sender As Object, e As TextChangedEventArgs)
        txtPlaceholderLibro.Visibility = If(String.IsNullOrEmpty(txtBuscarLibro.Text), Visibility.Visible, Visibility.Collapsed)

        Dim textoBusqueda As String = txtBuscarLibro.Text.Trim()

        If _cancellationTokenSourceLibros IsNot Nothing Then
            _cancellationTokenSourceLibros.Cancel()
            _cancellationTokenSourceLibros.Dispose()
        End If

        _cancellationTokenSourceLibros = New CancellationTokenSource()
        Dim token As CancellationToken = _cancellationTokenSourceLibros.Token

        Try
            Await Task.Delay(500, token)
            Await CargarLibrosAsync(textoBusqueda)
        Catch ex As TaskCanceledException
            ' Ignorado al seguir escribiendo
        Catch ex As Exception
            MessageBox.Show("Error al buscar libro: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Async Function CargarLibrosAsync(Optional filtro As String = "") As Task
        Try
            Dim libros = Await Task.Run(Function() _libroBLL.solicitarListaLibros(filtro))
            icLibros.ItemsSource = libros
            lblContadorLibros.Text = $"{libros.Count} libros registrados"
        Catch ex As Exception
            MessageBox.Show("Error al cargar libros: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Function

    Private Sub BtnSeleccionarLibro_Click(sender As Object, e As RoutedEventArgs)
        Dim btn As Button = TryCast(sender, Button)
        Dim libro As LibroItem = TryCast(btn?.Tag, LibroItem)

        If libro IsNot Nothing Then
            If libro.StockDisponible <= 0 Then
                MessageBox.Show($"El libro '{libro.Titulo}' no cuenta con ejemplares disponibles para préstamo.", "Sin stock disponible", MessageBoxButton.OK, MessageBoxImage.Warning)
                Return
            End If

            _libroSeleccionado = libro
            MostrarPaso3Plazo()
        End If
    End Sub

    ' =========================== PASO 3: SELECCIONAR PLAZO ===========================

    Private Sub MostrarPaso3Plazo()
        If _socioSeleccionado Is Nothing OrElse _libroSeleccionado Is Nothing Then
            MostrarPaso1Socios()
            Return
        End If

        ActualizarNavegacionLateral(esListaPrestamos:=False)

        ' Poblar vista previa con datos seleccionados
        lblPreviewLibroTitulo.Text = _libroSeleccionado.Titulo
        lblPreviewLibroDetalle.Text = $"{_libroSeleccionado.Autor} · ISBN {_libroSeleccionado.IsbnFormateado}"
        lblPreviewSocioNombre.Text = _socioSeleccionado.NombreCompleto
        lblPreviewSocioDetalle.Text = $"DNI {_socioSeleccionado.DniFormateado}"

        ' Seleccionar 14 días por defecto
        rbPlazo14.IsChecked = True
        ActualizarEstiloBordesPlazo()

        ' Cambiar vista
        gridListaPrestamos.Visibility = Visibility.Collapsed
        gridPaso1Socio.Visibility = Visibility.Collapsed
        gridPaso2Libro.Visibility = Visibility.Collapsed
        gridPasoExito.Visibility = Visibility.Collapsed
        gridPaso3Plazo.Visibility = Visibility.Visible
        scrollPrincipal.ScrollToTop()
    End Sub

    Private Sub BtnVolverPaso2_Click(sender As Object, e As RoutedEventArgs)
        MostrarPaso2Libros()
    End Sub

    Private Sub Plazo_Checked(sender As Object, e As RoutedEventArgs)
        ActualizarEstiloBordesPlazo()
    End Sub

    Private Sub BorderPlazo7_MouseDown(sender As Object, e As Input.MouseButtonEventArgs)
        rbPlazo7.IsChecked = True
    End Sub

    Private Sub BorderPlazo14_MouseDown(sender As Object, e As Input.MouseButtonEventArgs)
        rbPlazo14.IsChecked = True
    End Sub

    Private Sub BorderPlazo30_MouseDown(sender As Object, e As Input.MouseButtonEventArgs)
        rbPlazo30.IsChecked = True
    End Sub

    Private Sub ActualizarEstiloBordesPlazo()
        If borderPlazo7 Is Nothing OrElse borderPlazo14 Is Nothing OrElse borderPlazo30 Is Nothing Then
            Return
        End If

        Dim colorBorderUnselected As New SolidColorBrush(CType(ColorConverter.ConvertFromString("#E5E7EB"), Color))
        Dim colorBgUnselected As New SolidColorBrush(CType(ColorConverter.ConvertFromString("#FFFFFF"), Color))
        Dim colorBorderSelected As New SolidColorBrush(CType(ColorConverter.ConvertFromString("#1D3B31"), Color))
        Dim colorBgSelected As New SolidColorBrush(CType(ColorConverter.ConvertFromString("#EFF5F1"), Color))

        borderPlazo7.BorderBrush = If(rbPlazo7.IsChecked = True, colorBorderSelected, colorBorderUnselected)
        borderPlazo7.Background = If(rbPlazo7.IsChecked = True, colorBgSelected, colorBgUnselected)

        borderPlazo14.BorderBrush = If(rbPlazo14.IsChecked = True, colorBorderSelected, colorBorderUnselected)
        borderPlazo14.Background = If(rbPlazo14.IsChecked = True, colorBgSelected, colorBgUnselected)

        borderPlazo30.BorderBrush = If(rbPlazo30.IsChecked = True, colorBorderSelected, colorBorderUnselected)
        borderPlazo30.Background = If(rbPlazo30.IsChecked = True, colorBgSelected, colorBgUnselected)
    End Sub

    Private Async Sub BtnRegistrarPrestamoFinal_Click(sender As Object, e As RoutedEventArgs)
        If _socioSeleccionado Is Nothing OrElse _libroSeleccionado Is Nothing Then
            MessageBox.Show("Faltan datos para procesar el préstamo.", "Datos Incompletos", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim plazoDias As Integer = 14
        If rbPlazo7.IsChecked = True Then
            plazoDias = 7
        ElseIf rbPlazo30.IsChecked = True Then
            plazoDias = 30
        End If

        btnRegistrarPrestamoFinal.IsEnabled = False

        Try
            Dim prestamo As Prestamo = Await Task.Run(Function()
                Return _prestamoBLL.registrarPrestamo(_libroSeleccionado.IdLibro, _socioSeleccionado.IdSocio, plazoDias)
            End Function)

            ' Configurar panel de éxito
            lblExitoSocio.Text = $"{_socioSeleccionado.NombreCompleto} (DNI: {_socioSeleccionado.DniFormateado})"
            lblExitoLibro.Text = $"{_libroSeleccionado.Titulo} ({_libroSeleccionado.IsbnFormateado})"
            lblExitoFechaLimite.Text = $"{prestamo.fechaLimite:dd/MM/yyyy} ({plazoDias} días de plazo)"
            lblExitoMensaje.Text = $"Se ha registrado exitosamente el préstamo N° {prestamo.idPrestamo} a nombre de {_socioSeleccionado.NombreCompleto}."

            ' Mostrar pantalla de éxito
            gridListaPrestamos.Visibility = Visibility.Collapsed
            gridPaso1Socio.Visibility = Visibility.Collapsed
            gridPaso2Libro.Visibility = Visibility.Collapsed
            gridPaso3Plazo.Visibility = Visibility.Collapsed
            gridPasoExito.Visibility = Visibility.Visible
            scrollPrincipal.ScrollToTop()

        Catch ex As Exception
            MessageBox.Show("Error al registrar el préstamo: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            btnRegistrarPrestamoFinal.IsEnabled = True
        End Try
    End Sub

    ' =========================== PANTALLA DE ÉXITO ===========================

    Private Async Sub BtnVolverMenuPrestamos_Click(sender As Object, e As RoutedEventArgs)
        _socioSeleccionado = Nothing
        _libroSeleccionado = Nothing
        MostrarPaso1Socios()
        Await CargarSociosAsync()
    End Sub

    ' =========================== NAVEGACIÓN LATERAL Y GENERAL ===========================

    Private Sub ActualizarNavegacionLateral(esListaPrestamos As Boolean)
        Dim verdeActivo As New SolidColorBrush(CType(ColorConverter.ConvertFromString("#1D3B31"), Color))
        Dim verdeFondo As New SolidColorBrush(CType(ColorConverter.ConvertFromString("#EBF3EF"), Color))
        Dim grisTexto As New SolidColorBrush(CType(ColorConverter.ConvertFromString("#4B5563"), Color))

        If esListaPrestamos Then
            btnNavListaPrestamos.Background = verdeFondo
            textNavListaPrestamos.Foreground = verdeActivo
            textNavListaPrestamos.FontWeight = FontWeights.SemiBold
            iconNavListaPrestamos.Stroke = verdeActivo

            btnNavRegistrarPrestamo.Background = Brushes.Transparent
            textNavRegistrarPrestamo.Foreground = grisTexto
            textNavRegistrarPrestamo.FontWeight = FontWeights.Normal
            iconNavRegistrarPrestamo.Stroke = grisTexto
        Else
            btnNavRegistrarPrestamo.Background = verdeFondo
            textNavRegistrarPrestamo.Foreground = verdeActivo
            textNavRegistrarPrestamo.FontWeight = FontWeights.SemiBold
            iconNavRegistrarPrestamo.Stroke = verdeActivo

            btnNavListaPrestamos.Background = Brushes.Transparent
            textNavListaPrestamos.Foreground = grisTexto
            textNavListaPrestamos.FontWeight = FontWeights.Normal
            iconNavListaPrestamos.Stroke = grisTexto
        End If
    End Sub

    Private Sub BtnVolverMenu_Click(sender As Object, e As RoutedEventArgs)
        CerrarYVolverAlMenu()
    End Sub

    Private Sub PrestamosWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
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
