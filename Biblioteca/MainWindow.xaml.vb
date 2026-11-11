Class MainWindow

    Private Sub BtnSocios_Click(sender As Object, e As RoutedEventArgs)
        Dim ventanaSocios As New SociosWindow()
        ventanaSocios.Owner = Me
        ventanaSocios.Show()
        Me.Hide()
    End Sub

    Private Sub BtnLibros_Click(sender As Object, e As RoutedEventArgs)
        Dim ventanaLibros As New LibrosWindow()
        ventanaLibros.Owner = Me
        ventanaLibros.Show()
        Me.Hide()
    End Sub

    Private Sub BtnPrestamos_Click(sender As Object, e As RoutedEventArgs)
        Dim ventanaPrestamos As New PrestamosWindow()
        ventanaPrestamos.Owner = Me
        ventanaPrestamos.Show()
        Me.Hide()
    End Sub



End Class
