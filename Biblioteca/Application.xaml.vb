Class Application

    Protected Overrides Sub OnStartup(e As StartupEventArgs)
        MyBase.OnStartup(e)
        DatabaseHelper.InicializarBaseDatos()
    End Sub

End Class
