Imports Microsoft.Data.Sqlite

Public Module DatabaseHelper

    Public Const DbFileName As String = "biblioteca.db"
    Public ReadOnly ConnectionString As String = $"Data Source={DbFileName}"

    Public Sub InicializarBaseDatos()
        Using conexion As New SqliteConnection(ConnectionString)
            conexion.Open()

            Using comandoPragmas As SqliteCommand = conexion.CreateCommand()
                comandoPragmas.CommandText = "PRAGMA foreign_keys = ON;"
                comandoPragmas.ExecuteNonQuery()
            End Using

            Dim scriptTablas As String = "
                CREATE TABLE IF NOT EXISTS Socios (
                    idSocio INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombreCompleto TEXT NOT NULL,
                    dni INTEGER NOT NULL UNIQUE,
                    email TEXT,
                    telefono TEXT,
                    estado INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS Libros (
                    idLibro INTEGER PRIMARY KEY AUTOINCREMENT,
                    isbn INTEGER NOT NULL,
                    titulo TEXT NOT NULL,
                    autor TEXT NOT NULL,
                    editorial TEXT,
                    stockTotal INTEGER NOT NULL DEFAULT 0,
                    stockDisponible INTEGER NOT NULL DEFAULT 0
                );

                CREATE TABLE IF NOT EXISTS Prestamos (
                    idPrestamo INTEGER PRIMARY KEY AUTOINCREMENT,
                    idSocio INTEGER NOT NULL,
                    idLibro INTEGER NOT NULL,
                    fechaPrestamo TEXT NOT NULL,
                    fechaLimite TEXT NOT NULL,
                    fechaDevolucion TEXT,
                    fueDevuelto INTEGER NOT NULL DEFAULT 0,
                    FOREIGN KEY (idSocio) REFERENCES Socios (idSocio),
                    FOREIGN KEY (idLibro) REFERENCES Libros (idLibro)
                );
            "

            Using comando As SqliteCommand = conexion.CreateCommand()
                comando.CommandText = scriptTablas
                comando.ExecuteNonQuery()
            End Using

            ' Semilla de libros si la tabla está vacía
            Using cmdCount As SqliteCommand = conexion.CreateCommand()
                cmdCount.CommandText = "SELECT COUNT(*) FROM Libros;"
                Dim total As Long = Convert.ToInt64(cmdCount.ExecuteScalar())
                If total = 0 Then
                    Dim scriptSeedLibros As String = "
                        INSERT INTO Libros (isbn, titulo, autor, editorial, stockTotal, stockDisponible) VALUES
                        (9788437604947, 'Cien años de soledad', 'Gabriel García Márquez', 'Cátedra', 5, 5),
                        (9788478887194, 'El principito', 'Antoine de Saint-Exupéry', 'Salamandra', 3, 3),
                        (9788408043645, 'La sombra del viento', 'Carlos Ruiz Zafón', 'Planeta', 4, 4),
                        (9788420467284, 'Don Quijote de la Mancha', 'Miguel de Cervantes', 'Alfaguara', 2, 2),
                        (9788437604572, 'Rayuela', 'Julio Cortázar', 'Cátedra', 3, 3),
                        (9788497592208, 'La casa de los espíritus', 'Isabel Allende', 'Debolsillo', 4, 4),
                        (9788466319128, 'Ficciones', 'Jorge Luis Borges', 'Debolsillo', 5, 5);
                    "
                    Using cmdInsert As SqliteCommand = conexion.CreateCommand()
                        cmdInsert.CommandText = scriptSeedLibros
                        cmdInsert.ExecuteNonQuery()
                    End Using
                End If
            End Using

            ' Semilla de socios si la tabla está vacía
            Using cmdCountSocios As SqliteCommand = conexion.CreateCommand()
                cmdCountSocios.CommandText = "SELECT COUNT(*) FROM Socios;"
                Dim totalSocios As Long = Convert.ToInt64(cmdCountSocios.ExecuteScalar())
                If totalSocios = 0 Then
                    Dim scriptSeedSocios As String = "
                        INSERT INTO Socios (idSocio, nombreCompleto, dni, email, telefono, estado) VALUES
                        (1, 'María García López', 32456789, 'maria.garcia@example.com', '1145678901', 0),
                        (2, 'Juan Pérez Martín', 28734561, 'juan.perez@example.com', '1145678902', 0),
                        (3, 'Lucía Fernández Ruiz', 35891234, 'lucia.fernandez@example.com', '1145678903', 1),
                        (4, 'Carlos Sánchez Díaz', 30567892, 'carlos.sanchez@example.com', '1145678904', 0),
                        (5, 'Ana Martínez Torres', 33219876, 'ana.martinez@example.com', '1145678905', 0),
                        (6, 'Diego Romero Gil', 27843615, 'diego.romero@example.com', '1145678906', 1),
                        (7, 'Elena Navarro López', 36125478, 'elena.navarro@example.com', '1145678907', 0);
                    "
                    Using cmdInsertSocios As SqliteCommand = conexion.CreateCommand()
                        cmdInsertSocios.CommandText = scriptSeedSocios
                        cmdInsertSocios.ExecuteNonQuery()
                    End Using
                End If
            End Using
        End Using
    End Sub

End Module

