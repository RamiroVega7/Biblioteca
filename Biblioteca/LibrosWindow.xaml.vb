Imports System.Text.RegularExpressions
Imports System.Threading

Class LibrosWindow

    Private ReadOnly _libroBLL As New LibroBLL()
    Private _cancellationTokenModificar As CancellationTokenSource
    Private _cancellationTokenEliminar As CancellationTokenSource
    Private _libroSeleccionado As LibroItem = Nothing

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        MostrarModuloAgregarLibro()
    End Sub

    ' =========================== NAVEGACIÓN Y VISTAS ===========================

    ''' <summary>
    ''' Muestra la vista para agregar un nuevo libro al catálogo.
    ''' </summary>
    Public Sub MostrarModuloAgregarLibro()
        ActualizarNavegacionLateral(opcionSeleccionada:=0)

        gridAgregarLibro.Visibility = Visibility.Visible
        gridModificarLibro.Visibility = Visibility.Collapsed
        gridEliminarLibro.Visibility = Visibility.Collapsed
        scrollPrincipal.ScrollToTop()

        txtNuevoIsbn.Focus()
    End Sub

    ''' <summary>
    ''' Muestra la vista para modificar datos o stock de un libro existente.
    ''' Carga dinámicamente los libros desde la base de datos a través de la BLL.
    ''' </summary>
    Public Async Sub MostrarModuloModificarLibro()
        ActualizarNavegacionLateral(opcionSeleccionada:=1)

        gridAgregarLibro.Visibility = Visibility.Collapsed
        gridModificarLibro.Visibility = Visibility.Visible
        gridEliminarLibro.Visibility = Visibility.Collapsed
        scrollPrincipal.ScrollToTop()

        txtBuscarModificar.Focus()

        Await CargarListaLibrosModificarAsync(txtBuscarModificar.Text.Trim())
    End Sub

    ''' <summary>
    ''' Muestra la vista para dar de baja o eliminar un libro del catálogo.
    ''' Carga de forma asíncrona la lista de libros disponibles.
    ''' </summary>
    Public Async Sub MostrarModuloEliminarLibro()
        ActualizarNavegacionLateral(opcionSeleccionada:=2)

        gridAgregarLibro.Visibility = Visibility.Collapsed
        gridModificarLibro.Visibility = Visibility.Collapsed
        gridEliminarLibro.Visibility = Visibility.Visible
        scrollPrincipal.ScrollToTop()

        txtBuscarEliminar.Focus()

        Await CargarListaLibrosEliminarAsync(txtBuscarEliminar.Text.Trim())
    End Sub

    Private Sub BtnNavAgregarLibro_Click(sender As Object, e As RoutedEventArgs)
        MostrarModuloAgregarLibro()
    End Sub

    Private Sub BtnNavModificarLibro_Click(sender As Object, e As RoutedEventArgs)
        MostrarModuloModificarLibro()
    End Sub

    Private Sub BtnNavEliminarLibro_Click(sender As Object, e As RoutedEventArgs)
        MostrarModuloEliminarLibro()
    End Sub

    ''' <summary>
    ''' Actualiza los estilos visuales de las opciones de la barra de navegación lateral.
    ''' 0 = Agregar libro, 1 = Modificar libro, 2 = Eliminar libro
    ''' </summary>
    Private Sub ActualizarNavegacionLateral(opcionSeleccionada As Integer)
        Dim verdeActivo As Brush = New SolidColorBrush(Color.FromRgb(&H1D, &H3B, &H31))
        Dim verdeFondoPill As Brush = New SolidColorBrush(Color.FromRgb(&HEB, &HF3, &HEF))
        Dim grisTexto As Brush = New SolidColorBrush(Color.FromRgb(&H4B, &H55, &H63))

        ' Opción 0: Agregar libro
        If opcionSeleccionada = 0 Then
            btnNavAgregarLibro.Background = verdeFondoPill
            textNavAgregarLibro.Foreground = verdeActivo
            textNavAgregarLibro.FontWeight = FontWeights.SemiBold
            iconNavAgregarLibro.Stroke = verdeActivo
        Else
            btnNavAgregarLibro.Background = Brushes.Transparent
            textNavAgregarLibro.Foreground = grisTexto
            textNavAgregarLibro.FontWeight = FontWeights.Normal
            iconNavAgregarLibro.Stroke = grisTexto
        End If

        ' Opción 1: Modificar libro
        If opcionSeleccionada = 1 Then
            btnNavModificarLibro.Background = verdeFondoPill
            textNavModificarLibro.Foreground = verdeActivo
            textNavModificarLibro.FontWeight = FontWeights.SemiBold
            iconNavModificarLibro.Stroke = verdeActivo
        Else
            btnNavModificarLibro.Background = Brushes.Transparent
            textNavModificarLibro.Foreground = grisTexto
            textNavModificarLibro.FontWeight = FontWeights.Normal
            iconNavModificarLibro.Stroke = grisTexto
        End If

        ' Opción 2: Eliminar libro
        If opcionSeleccionada = 2 Then
            btnNavEliminarLibro.Background = verdeFondoPill
            textNavEliminarLibro.Foreground = verdeActivo
            textNavEliminarLibro.FontWeight = FontWeights.SemiBold
            iconNavEliminarLibro.Stroke = verdeActivo
        Else
            btnNavEliminarLibro.Background = Brushes.Transparent
            textNavEliminarLibro.Foreground = grisTexto
            textNavEliminarLibro.FontWeight = FontWeights.Normal
            iconNavEliminarLibro.Stroke = grisTexto
        End If
    End Sub

    ' =========================== GESTIÓN Y CARGA ASÍNCRONA DE LIBROS (MODIFICAR) ===========================

    ''' <summary>
    ''' Carga de forma asíncrona los libros en la tabla de modificación aplicando filtros.
    ''' </summary>
    Public Async Function CargarListaLibrosModificarAsync(Optional filtro As String = "") As Task
        Try
            Dim libros As List(Of LibroItem) = Await Task.Run(Function()
                Return _libroBLL.solicitarListaLibros(filtro)
            End Function)

            icListaLibrosModificar.ItemsSource = libros

            Dim total As Integer = libros.Count
            lblContadorLibrosModificar.Text = If(total = 1, "1 libro en catálogo", $"{total} libros en catálogo")

            panelSinResultadosModificar.Visibility = If(total = 0, Visibility.Visible, Visibility.Collapsed)

            ' Si el libro actualmente seleccionado sigue en la lista, actualizar la referencia
            If _libroSeleccionado IsNot Nothing Then
                Dim libroActualizado As LibroItem = libros.FirstOrDefault(Function(l) l.IdLibro = _libroSeleccionado.IdLibro)
                If libroActualizado IsNot Nothing Then
                    _libroSeleccionado = libroActualizado
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error al cargar el catálogo de libros: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Function

    ''' <summary>
    ''' Búsqueda con debounce (400ms) para la tabla de modificación de libros.
    ''' </summary>
    Private Async Sub TxtBuscarModificar_TextChanged(sender As Object, e As TextChangedEventArgs)
        If txtPlaceholderBuscarModificar IsNot Nothing Then
            txtPlaceholderBuscarModificar.Visibility = If(String.IsNullOrEmpty(txtBuscarModificar.Text), Visibility.Visible, Visibility.Collapsed)
        End If

        Dim textoBusqueda As String = txtBuscarModificar.Text.Trim()

        If _cancellationTokenModificar IsNot Nothing Then
            _cancellationTokenModificar.Cancel()
            _cancellationTokenModificar.Dispose()
        End If

        _cancellationTokenModificar = New CancellationTokenSource()
        Dim token As CancellationToken = _cancellationTokenModificar.Token

        Try
            Await Task.Delay(400, token)
            Await CargarListaLibrosModificarAsync(textoBusqueda)
        Catch ex As TaskCanceledException
            ' Cancelado por nuevo tecleo del usuario
        End Try
    End Sub

    ' =========================== GESTIÓN Y CARGA ASÍNCRONA DE LIBROS (ELIMINAR) ===========================

    ''' <summary>
    ''' Carga de forma asíncrona los libros en la tabla de eliminación aplicando filtros.
    ''' </summary>
    Public Async Function CargarListaLibrosEliminarAsync(Optional filtro As String = "") As Task
        Try
            Dim libros As List(Of LibroItem) = Await Task.Run(Function()
                Return _libroBLL.solicitarListaLibros(filtro)
            End Function)

            icListaLibrosEliminar.ItemsSource = libros

            Dim total As Integer = libros.Count
            lblContadorLibrosEliminar.Text = If(total = 1, "1 libro en catálogo", $"{total} libros en catálogo")

            panelSinResultadosEliminar.Visibility = If(total = 0, Visibility.Visible, Visibility.Collapsed)

        Catch ex As Exception
            MessageBox.Show("Error al cargar los libros para eliminar: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Function

    ''' <summary>
    ''' Búsqueda con debounce (400ms) para la tabla de eliminación de libros.
    ''' </summary>
    Private Async Sub TxtBuscarEliminar_TextChanged(sender As Object, e As TextChangedEventArgs)
        If txtPlaceholderBuscarEliminar IsNot Nothing Then
            txtPlaceholderBuscarEliminar.Visibility = If(String.IsNullOrEmpty(txtBuscarEliminar.Text), Visibility.Visible, Visibility.Collapsed)
        End If

        Dim textoBusqueda As String = txtBuscarEliminar.Text.Trim()

        If _cancellationTokenEliminar IsNot Nothing Then
            _cancellationTokenEliminar.Cancel()
            _cancellationTokenEliminar.Dispose()
        End If

        _cancellationTokenEliminar = New CancellationTokenSource()
        Dim token As CancellationToken = _cancellationTokenEliminar.Token

        Try
            Await Task.Delay(400, token)
            Await CargarListaLibrosEliminarAsync(textoBusqueda)
        Catch ex As TaskCanceledException
            ' Cancelado por nuevo tecleo del usuario
        End Try
    End Sub

    ''' <summary>
    ''' Procesa la eliminación permanente de un libro del catálogo tras confirmación del usuario.
    ''' </summary>
    Private Async Sub BtnEliminarLibro_Click(sender As Object, e As RoutedEventArgs)
        Dim boton As Button = TryCast(sender, Button)
        Dim item As LibroItem = TryCast(boton?.Tag, LibroItem)

        If item Is Nothing Then Return

        Dim confirmacion = MessageBox.Show(
            $"¿Estás seguro de que deseas eliminar permanentemente el libro '{item.Titulo}' (ISBN: {item.IsbnFormateado}) del catálogo?{Environment.NewLine}{Environment.NewLine}Esta acción no se puede deshacer.",
            "Confirmar Eliminación",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question)

        If confirmacion <> MessageBoxResult.Yes Then Return

        Try
            Dim idLibro As Integer = item.IdLibro
            Dim eliminado As Boolean = Await Task.Run(Function()
                Return _libroBLL.eliminarLibro(idLibro)
            End Function)

            If eliminado Then
                MessageBox.Show($"El libro '{item.Titulo}' fue eliminado exitosamente del catálogo.",
                                "Eliminación Exitosa", MessageBoxButton.OK, MessageBoxImage.Information)

                ' Refrescar listado de libros para eliminar
                Await CargarListaLibrosEliminarAsync(txtBuscarEliminar.Text.Trim())

                ' Si el libro eliminado era el seleccionado en la vista Modificar, resetearlo
                If _libroSeleccionado IsNot Nothing AndAlso _libroSeleccionado.IdLibro = idLibro Then
                    LimpiarFormularioModificar()
                End If
            End If

        Catch ex As InvalidOperationException
            MessageBox.Show(ex.Message, "Regla de Negocio", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error al eliminar el libro: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    ' =========================== ACCIONES DEL FORMULARIO DE ALTA (AGREGAR) ===========================

    ''' <summary>
    ''' Valida los datos e incorpora un nuevo libro al catálogo a través de la capa BLL.
    ''' </summary>
    Private Async Sub BtnGuardarNuevoLibro_Click(sender As Object, e As RoutedEventArgs)
        Dim isbnTexto As String = txtNuevoIsbn.Text.Trim()
        Dim titulo As String = txtNuevoTitulo.Text.Trim()
        Dim autor As String = txtNuevoAutor.Text.Trim()
        Dim editorial As String = txtNuevaEditorial.Text.Trim()
        Dim stockTotalTexto As String = txtNuevoStockTotal.Text.Trim()

        btnGuardarNuevoLibro.IsEnabled = False

        Try
            Dim nuevoLibro As Libro = Await Task.Run(Function()
                Return _libroBLL.registrarLibro(isbnTexto, titulo, autor, editorial, stockTotalTexto)
            End Function)

            MessageBox.Show($"El libro '{nuevoLibro.titulo}' fue incorporado exitosamente al catálogo.",
                            "Registro Exitoso", MessageBoxButton.OK, MessageBoxImage.Information)

            LimpiarFormularioNuevoLibro()

        Catch ex As ArgumentException
            MessageBox.Show(ex.Message, "Validación de Datos", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As InvalidOperationException
            MessageBox.Show(ex.Message, "Regla de Negocio", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error al guardar el libro: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            btnGuardarNuevoLibro.IsEnabled = True
        End Try
    End Sub

    ''' <summary>
    ''' Limpia los campos del formulario de nuevo libro y restablece los valores predeterminados.
    ''' </summary>
    Private Sub BtnLimpiarNuevoLibro_Click(sender As Object, e As RoutedEventArgs)
        LimpiarFormularioNuevoLibro()
    End Sub

    Private Sub LimpiarFormularioNuevoLibro()
        txtNuevoIsbn.Clear()
        txtNuevoTitulo.Clear()
        txtNuevoAutor.Clear()
        txtNuevaEditorial.Clear()
        txtNuevoStockTotal.Text = "1"
        txtNuevoStockDisponible.Text = "1"
        txtNuevoIsbn.Focus()
    End Sub

    ' =========================== ACCIONES DEL FORMULARIO DE MODIFICACIÓN ===========================

    ''' <summary>
    ''' Carga el libro seleccionado en el formulario de edición.
    ''' </summary>
    Private Sub BtnSeleccionarLibroModificar_Click(sender As Object, e As RoutedEventArgs)
        Dim boton As Button = TryCast(sender, Button)
        Dim item As LibroItem = TryCast(boton?.Tag, LibroItem)

        If item Is Nothing Then Return

        _libroSeleccionado = item

        txtModificarIsbn.Text = item.Isbn.ToString()
        txtModificarTitulo.Text = item.Titulo
        txtModificarAutor.Text = item.Autor
        txtModificarEditorial.Text = If(item.Editorial, "")
        txtModificarStockTotal.Text = item.StockTotal.ToString()
        txtModificarStockDisponible.Text = item.StockDisponible.ToString()

        lblEstadoSeleccionModificar.Text = $"(Editando: #{item.IdLibro} - {item.Titulo})"
        lblEstadoSeleccionModificar.Foreground = New SolidColorBrush(Color.FromRgb(&H1D, &H3B, &H31))
        lblAyudaFormularioModificar.Text = $"Modifica los campos del libro y haz clic en 'Guardar Cambios' para persistir en la base de datos."

        txtModificarTitulo.Focus()
        txtModificarTitulo.SelectAll()
    End Sub

    ''' <summary>
    ''' Valida los datos y persiste las modificaciones a través de la capa BLL.
    ''' </summary>
    Private Async Sub BtnGuardarModificar_Click(sender As Object, e As RoutedEventArgs)
        If _libroSeleccionado Is Nothing Then
            MessageBox.Show("Por favor, selecciona primero un libro de la lista superior para editar sus datos.",
                            "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim idLibro As Integer = _libroSeleccionado.IdLibro
        Dim isbnTexto As String = txtModificarIsbn.Text.Trim()
        Dim titulo As String = txtModificarTitulo.Text.Trim()
        Dim autor As String = txtModificarAutor.Text.Trim()
        Dim editorial As String = txtModificarEditorial.Text.Trim()
        Dim stockTotalTexto As String = txtModificarStockTotal.Text.Trim()

        btnGuardarModificar.IsEnabled = False

        Try
            Dim libroActualizado As Libro = Await Task.Run(Function()
                Return _libroBLL.modificarLibro(idLibro, isbnTexto, titulo, autor, editorial, stockTotalTexto)
            End Function)

            MessageBox.Show($"El libro '{libroActualizado.titulo}' fue actualizado con éxito en el catálogo.",
                            "Modificación Exitosa", MessageBoxButton.OK, MessageBoxImage.Information)

            ' Recargar lista desde la base de datos para reflejar los cambios
            Await CargarListaLibrosModificarAsync(txtBuscarModificar.Text.Trim())

            ' Actualizar campos y etiquetas del formulario
            txtModificarStockDisponible.Text = libroActualizado.stockDisponible.ToString()
            txtModificarStockTotal.Text = libroActualizado.stockTotal.ToString()
            lblEstadoSeleccionModificar.Text = $"(Guardado: #{libroActualizado.idLibro} - {libroActualizado.titulo})"
            lblAyudaFormularioModificar.Text = $"Cambios guardados. Stock disponible recalculado a {libroActualizado.stockDisponible} ejemplares."

        Catch ex As ArgumentException
            MessageBox.Show(ex.Message, "Validación de Datos", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As InvalidOperationException
            MessageBox.Show(ex.Message, "Regla de Negocio", MessageBoxButton.OK, MessageBoxImage.Warning)
        Catch ex As Exception
            MessageBox.Show("Ocurrió un error al guardar los cambios: " & ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        Finally
            btnGuardarModificar.IsEnabled = True
        End Try
    End Sub

    ''' <summary>
    ''' Cancela la edición y limpia los campos del formulario.
    ''' </summary>
    Private Sub BtnCancelarModificar_Click(sender As Object, e As RoutedEventArgs)
        LimpiarFormularioModificar()
    End Sub

    Private Sub LimpiarFormularioModificar()
        _libroSeleccionado = Nothing
        txtModificarIsbn.Clear()
        txtModificarTitulo.Clear()
        txtModificarAutor.Clear()
        txtModificarEditorial.Clear()
        txtModificarStockTotal.Clear()
        txtModificarStockDisponible.Clear()

        lblEstadoSeleccionModificar.Text = "(Ningún libro seleccionado)"
        lblEstadoSeleccionModificar.Foreground = New SolidColorBrush(Color.FromRgb(&H6B, &H72, &H80))
        lblAyudaFormularioModificar.Text = "Selecciona un libro de la lista superior para cargar y modificar sus datos."
    End Sub

    ''' <summary>
    ''' Máscara para admitir exclusivamente caracteres numéricos en campos de ISBN y stock.
    ''' </summary>
    Private Sub TxtSoloNumeros_PreviewTextInput(sender As Object, e As TextCompositionEventArgs)
        Dim regex As New Regex("[^0-9]+")
        e.Handled = regex.IsMatch(e.Text)
    End Sub

    ' =========================== NAVEGACIÓN Y CIERRE ===========================

    Private Sub BtnVolverMenu_Click(sender As Object, e As RoutedEventArgs)
        CerrarYVolverAlMenu()
    End Sub

    Private Sub LibrosWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
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
