# Contexto del Proyecto: Biblioteca

Sistema de escritorio para la administración integral de una biblioteca (gestión de socios, catálogo de libros, préstamos y devoluciones).

---

## 1. Información General y Tecnologías

- **Nombre del Proyecto:** Biblioteca - Sistema de Gestión
- **Tipo de Aplicación:** Aplicación de Escritorio (WPF - Windows Presentation Foundation)
- **Lenguaje Principal:** Visual Basic .NET (`VB.NET`) — *Regla estricta: todo el código fuente debe ser `.vb`*
- **Framework de Destino:** `.NET 10.0` (`net10.0-windows`)
- **Gestor de Base de Datos:** SQLite (`Microsoft.Data.Sqlite` v10.0.14)
- **Archivo de Base de Datos Local:** `biblioteca.db` (ubicado en el directorio de ejecución / raíz del proyecto)
- **Solución / Proyecto:** `Biblioteca.slnx`, `Biblioteca/Biblioteca.vbproj`

---

## 2. Arquitectura del Sistema (Arquitectura en Capas)

El proyecto implementa un patrón estricto de **Arquitectura en Capas**, garantizando una clara separación de responsabilidades, alta cohesión y bajo acoplamiento:

```
┌────────────────────────────────────────────────────────┐
│               UI (User Interface)                      │
│   Vistas WPF (XAML + Code-Behind VB.NET) & ViewModels │
└───────────────────────────┬────────────────────────────┘
                            │ Invoca
                            ▼
┌────────────────────────────────────────────────────────┐
│            BLL (Business Logic Layer)                  │
│       Reglas de Negocio, Validaciones y Formato        │
└───────────────────────────┬────────────────────────────┘
                            │ Invoca
                            ▼
┌────────────────────────────────────────────────────────┐
│             DAL (Data Access Layer)                    │
│      Acceso a Datos, Consultas SQL Parametrizadas      │
└───────────────────────────┬────────────────────────────┘
                            │ Lee / Escribe
                            ▼
┌────────────────────────────────────────────────────────┐
│             Base de Datos (SQLite)                     │
└────────────────────────────────────────────────────────┘

     ▲ Transversal a todas las capas:
     └─► BE (Business Entities) / Modelos del Dominio
```

---

## 3. Catálogo Detallado de Clases, Métodos y Reutilización

A continuación se detalla exhaustivamente cada capa, indicando qué clases contiene, qué métodos implementa cada una, su comportamiento interno y cómo reutilizarlos en futuros módulos.

---

### 3.1. Capa de Entidades de Negocio (BE - Business Entities)
*Ubicación:* `Biblioteca/Models/`  
Representa los modelos del dominio y las estructuras de datos fundamentales que fluyen entre todas las capas.

#### `Socio.vb`
- **`EstadoSocio` (Enum):**
  - `Activo = 0`: Socio habilitado para solicitar préstamos y operar.
  - `Inactivo = 1`: Socio deshabilitado (baja lógica).
- **`Socio` (Clase):**
  - *Propiedades:*
    - `idSocio As Integer`: Clave primaria autoincremental en SQLite.
    - `nombreCompleto As String`: Nombre y apellido del socio.
    - `dni As Long`: Documento de identidad (único en el sistema).
    - `email As String`: Dirección de correo electrónico (puede ser Nothing o DBNull).
    - `telefono As String`: Número telefónico de contacto (puede ser Nothing o DBNull).
    - `estado As EstadoSocio`: Estado actual (por defecto `EstadoSocio.Activo`).
  - *Guía de Reutilización:* Entidad requerida para asociar a los préstamos, para el futuro módulo ABM de administración de socios y para reportes de actividad por socio.

#### `Libro.vb`
- **`Libro` (Clase):**
  - *Propiedades:*
    - `idLibro As Integer`: Clave primaria autoincremental en SQLite.
    - `isbn As Long`: Código numérico internacional del libro (13 dígitos).
    - `titulo As String`: Título de la obra.
    - `autor As String`: Autor de la publicación.
    - `editorial As String`: Editorial responsable de la impresión.
    - `stockTotal As Integer`: Cantidad física total de ejemplares que posee la biblioteca.
    - `stockDisponible As Integer`: Ejemplares disponibles en estantería listos para ser prestados.
  - *Guía de Reutilización:* Base fundamental para el catálogo, control de inventario físico, validación de disponibilidad en préstamos y futuro módulo de gestión de libros.

#### `Prestamo.vb`
- **`Prestamo` (Clase):**
  - *Propiedades:*
    - `idPrestamo As Integer`: Clave primaria autoincremental en SQLite.
    - `socio As Socio`: Referencia completa a la entidad `Socio` prestataria.
    - `libro As Libro`: Referencia completa a la entidad `Libro` prestada.
    - `fechaPrestamo As Date`: Fecha en la que se otorga el préstamo.
    - `fechaLimite As Date`: Fecha de vencimiento pactada para la devolución.
    - `fechaDevolucion As Date?`: Fecha real en la que el libro fue reintegrado (`Nothing` si sigue prestado).
    - `fueDevuelto As Boolean`: Indicador booleano (`True` si ya se devolvió, `False` si sigue en curso).
  - *Guía de Reutilización:* Modelo central de transacciones de la biblioteca. Reutilizable para historial, auditoría, cálculo de multas/sanciones y reportes de morosidad.

---

### 3.2. Capa de Acceso a Datos (DAL - Data Access Layer)
*Ubicación:* `Biblioteca/DAL/` y `DatabaseHelper.vb`  
Responsable exclusivo de interactuar con SQLite mediante sentencias SQL parametrizadas, gestión de conexiones y control de transacciones.

#### `DatabaseHelper.vb` (Módulo)
- **Constantes y Campos:**
  - `DbFileName As String = "biblioteca.db"`: Nombre del archivo de base de datos local.
  - `ConnectionString As String`: Cadena de conexión SQLite configurada (`Data Source=biblioteca.db`).
- **Métodos:**
  - **`InicializarBaseDatos()`**
    - *Firma:* `Public Sub InicializarBaseDatos()`
    - *Qué hace:* Abre una conexión SQLite, activa la integridad referencial (`PRAGMA foreign_keys = ON;`), ejecuta el DDL de creación de tablas (`Socios`, `Libros`, `Prestamos`) con `CREATE TABLE IF NOT EXISTS` e inserta la semilla inicial de 7 libros si la tabla está vacía.
    - *Reutilización:* Se invoca en `Application.xaml.vb` al iniciar la app. Permite agregar scripts de migración o nuevas tablas (`Multas`, `Categorías`, etc.) en un único punto centralizado.

#### `SocioDAL.vb` (Clase)
- **Métodos:**
  - **`buscarSociosHabilitados(Optional filtro As String = "") As List(Of Socio)`**
    - *Firma:* `Public Function buscarSociosHabilitados(Optional filtro As String = "") As List(Of Socio)`
    - *Qué hace:* Consulta la tabla `Socios` filtrando exclusivamente socios activos (`WHERE estado = 0`). Si se provee un filtro, aplica `LIKE` insensible sobre `CAST(dni AS TEXT)` o `nombreCompleto`, ordenando alfabéticamente. Utiliza parámetros SQL seguros (`@filtro`) y delega la ejecución en `EjecutarConsultaSocios` y el mapeo en `MapearSocio`.
    - *Reutilización:* Búsqueda general de socios activos, selección en préstamos y consultas operativas. *(Nota: Mantiene `buscarSocios` como alias obsoleto por compatibilidad).*
  - **`buscarTodosLosSocios(Optional filtro As String = "") As List(Of Socio)`**
    - *Firma:* `Public Function buscarTodosLosSocios(Optional filtro As String = "") As List(Of Socio)`
    - *Qué hace:* Consulta todos los registros de la tabla `Socios` (activos e inactivos). Si se ingresa un filtro, evalúa con `LIKE` insensible sobre `CAST(dni AS TEXT)`, `nombreCompleto` o `email`. Ordena por `idSocio ASC` delegando en `EjecutarConsultaSocios`.
    - *Reutilización:* Grilla de administración de socios y gestión de altas, bajas lógicas y reactivaciones.
  - **`buscarTuplasSociosHabilitados(Optional filtro As String = "") As List(Of (id As Integer, nombreCompleto As String, dni As Long))`**
    - *Firma:* `Public Function buscarTuplasSociosHabilitados(Optional filtro As String = "") As List(Of (id As Integer, nombreCompleto As String, dni As Long))`
    - *Qué hace:* Ejecuta una consulta optimizada a nivel SQLite (`SELECT idSocio, nombreCompleto, dni FROM Socios WHERE estado = 0...`) extrayendo únicamente las 3 columnas necesarias para controles ligeros sin sobrecarga de memoria.
    - *Reutilización:* Ideal para poblar controles livianos como `ComboBox`, sugerencias y autocompletado rápido. *(Nota: Mantiene `buscarListaSocios` como alias obsoleto por compatibilidad).*
  - **`buscarSocioPorId(idSocio As Integer) As Socio`**
    - *Firma:* `Public Function buscarSocioPorId(idSocio As Integer) As Socio`
    - *Qué hace:* Consulta un único socio por su clave primaria (`WHERE idSocio = @idSocio LIMIT 1;`) y lo mapea de forma centralizada con `MapearSocio`.
    - *Reutilización:* Validación puntual de prestatarios en préstamos y búsqueda directa por identificador.
  - **`cambiarEstadoSocio(idSocio As Integer, nuevoEstado As EstadoSocio) As Boolean`**
    - *Firma:* `Public Function cambiarEstadoSocio(idSocio As Integer, nuevoEstado As EstadoSocio) As Boolean`
    - *Qué hace:* Actualiza el campo `estado` (`0` para Activo, `1` para Inactivo) del socio especificado mediante consulta SQL parametrizada.
  - **`guardarSocio(socio As Socio) As Integer`**
    - *Firma:* `Public Function guardarSocio(socio As Socio) As Integer`
    - *Qué hace:* Inserta un nuevo socio con `INSERT INTO Socios (...)` y obtiene la clave primaria generada con `last_insert_rowid()`.
  - **`existeDni(dni As Long, Optional idExcluir As Integer = 0) As Boolean`**
    - *Firma:* `Public Function existeDni(dni As Long, Optional idExcluir As Integer = 0) As Boolean`
    - *Qué hace:* Verifica la existencia previa de un DNI en la base de datos para prevenir registros duplicados.

#### `LibroDAL.vb` (Clase)
- **Métodos:**
  - **`buscarLibros(Optional filtro As String = "") As List(Of Libro)`**
    - *Firma:* `Public Function buscarLibros(Optional filtro As String = "") As List(Of Libro)`
    - *Qué hace:* Consulta la tabla `Libros`. Si se especifica `filtro`, busca coincidencias mediante `LIKE @filtro` sobre `isbn`, `titulo` o `autor`. Retorna todas las columnas mapeadas a entidades `Libro` con sus stocks actuales.
    - *Reutilización:* Búsqueda de libros en el catálogo general, pantalla de selección de libros en préstamos y módulo de administración y catálogo bibliográfico.
  - **`buscarLibroPorId(idLibro As Integer) As Libro`**
    - *Firma:* `Public Function buscarLibroPorId(idLibro As Integer) As Libro`
    - *Qué hace:* Consulta un único libro por su clave primaria (`WHERE idLibro = @idLibro LIMIT 1;`).
    - *Reutilización:* Búsqueda y validación puntual de libros para edición, verificación de existencia o detalle.
  - **`actualizarLibro(libro As Libro) As Boolean`**
    - *Firma:* `Public Function actualizarLibro(libro As Libro) As Boolean`
    - *Qué hace:* Actualiza los campos `isbn`, `titulo`, `autor`, `editorial`, `stockTotal` y `stockDisponible` mediante SQL parametrizado.
    - *Reutilización:* Edición y actualización de existencias en el módulo de administración de libros.
  - **`existeIsbn(isbn As Long, Optional idExcluir As Integer = 0) As Boolean`**
    - *Firma:* `Public Function existeIsbn(isbn As Long, Optional idExcluir As Integer = 0) As Boolean`
    - *Qué hace:* Verifica si ya existe otro libro con el mismo ISBN numérico en el catálogo (excluyendo opcionalmente el ID actual).
  - **`contarPrestamosActivos(idLibro As Integer) As Integer`**
    - *Firma:* `Public Function contarPrestamosActivos(idLibro As Integer) As Integer`
    - *Qué hace:* Consulta `SELECT COUNT(*) FROM Prestamos WHERE idLibro = @idLibro AND fueDevuelto = 0;` para evitar inconsistencias de inventario físico.
  - **`guardarLibro(libro As Libro) As Integer`**
    - *Firma:* `Public Function guardarLibro(libro As Libro) As Integer`
    - *Qué hace:* Inserta un nuevo registro en `Libros` (`isbn`, `titulo`, `autor`, `editorial`, `stockTotal`, `stockDisponible`) y recupera la clave primaria generada con `last_insert_rowid()`.
  - **`eliminarLibro(idLibro As Integer) As Boolean`**
    - *Firma:* `Public Function eliminarLibro(idLibro As Integer) As Boolean`
    - *Qué hace:* Ejecuta una transacción atómica: valida que no existan préstamos activos (`fueDevuelto = 0`), remueve préstamos históricos devueltos asociados (`fueDevuelto = 1`) para satisfacer la clave foránea, elimina la fila en `Libros` y confirma con `Commit()`.
  - **`buscarListaLibrosDisponibles() As List(Of (id As Integer, isbn As Long, titulo As String, autor As String))`**
    - *Firma:* `Public Function buscarListaLibrosDisponibles() As List(Of (id As Integer, isbn As Long, titulo As String, autor As String))`
    - *Qué hace:* Ejecuta `SELECT idLibro, isbn, titulo, autor FROM Libros WHERE stockDisponible > 0 ORDER BY titulo ASC;` y retorna tuplas con los libros que tienen disponibilidad inmediata.
    - *Reutilización:* Carga rápida de listas desplegables donde sólo se permite seleccionar libros con stock disponible para préstamo.

#### `PrestamoDAL.vb` (Clase)
- **Métodos:**
  - **`guardarPrestamo(prestamo As Prestamo) As Integer`**
    - *Firma:* `Public Function guardarPrestamo(prestamo As Prestamo) As Integer`
    - *Qué hace:* Ejecuta una **transacción atómica** en SQLite:
      1. Inserta el registro en `Prestamos` (`idSocio`, `idLibro`, `fechaPrestamo`, `fechaLimite`, `fechaDevolucion`, `fueDevuelto`) guardando fechas en formato ISO `yyyy-MM-dd`.
      2. Obtiene el ID asignado con `SELECT last_insert_rowid();` y lo actualiza en `prestamo.idPrestamo`.
      3. Descuenta automáticamente 1 unidad del stock disponible del libro: `UPDATE Libros SET stockDisponible = stockDisponible - 1 WHERE idLibro = @idLibro AND stockDisponible > 0;`.
      4. Si ambas operaciones son exitosas realiza `Commit()` y retorna el ID; si algo falla ejecuta `Rollback()` y relanza la excepción.
    - *Reutilización:* Punto único para registrar y persistir cualquier préstamo garantizando la consistencia del inventario.
  - **`registrarDevolucion(idPrestamo As Integer) As Boolean`**
    - *Firma:* `Public Function registrarDevolucion(idPrestamo As Integer) As Boolean`
    - *Qué hace:* Ejecuta una **transacción atómica** en SQLite:
      1. Verifica la existencia del préstamo y si ya fue devuelto (`fueDevuelto = 1`). Si ya estaba devuelto, retorna `True` sin alterar el stock.
      2. Si está pendiente, actualiza el préstamo: `UPDATE Prestamos SET fueDevuelto = 1, fechaDevolucion = @fechaDevolucion WHERE idPrestamo = @idPrestamo;` con la fecha actual.
      3. Reintegra la unidad al stock disponible del libro: `UPDATE Libros SET stockDisponible = stockDisponible + 1 WHERE idLibro = @idLibro;`.
      4. Confirma con `Commit()` y retorna `True`. En caso de fallo ejecuta `Rollback()`.
    - *Reutilización:* Se utiliza en la columna de acción de la lista de préstamos y es directamente reutilizable para futuros módulos de devolución mediante lectores de código de barras o devoluciones en lote.
  - **`buscarPrestamos(Optional filtro As String = "") As List(Of Prestamo)`**
    - *Firma:* `Public Function buscarPrestamos(Optional filtro As String = "") As List(Of Prestamo)`
    - *Qué hace:* Realiza un `INNER JOIN` entre `Prestamos`, `Socios` y `Libros`. Si se pasa `filtro`, busca por nombre de socio, título de libro o autor. Parsea las fechas ISO y reconstruye los objetos de dominio completos (`Prestamo` con sus referencias `Socio` y `Libro`).
    - *Reutilización:* Grilla general de préstamos, historial de préstamos de un socio en particular, historial de uso de un libro específico y reportes de morosidad.

---

### 3.3. Capa de Lógica de Negocio (BLL - Business Logic Layer)
*Ubicación:* `Biblioteca/BLL/`  
Aplica validaciones de negocio, transformaciones a modelos de vista (ViewModels), formateo cultural y orquesta las llamadas a la DAL.

#### `SocioBLL.vb` (Clase)
- **Campos:**
  - `_socioDAL As New SocioDAL()`: Instancia de acceso a datos de socios.
  - `_culturaEs As New CultureInfo("es-AR")`: Configuración regional argentina para números y fechas.
- **Métodos:**
  - **`solicitarSociosHabilitados(Optional filtro As String = "") As List(Of SocioItem)`**
    - *Firma:* `Public Function solicitarSociosHabilitados(Optional filtro As String = "") As List(Of SocioItem)`
    - *Qué hace:* Llama a `_socioDAL.buscarSocios(filtro)` y proyecta cada socio habilitado en un `SocioItem` mediante el método privado de mapeo `MapearListaSociosItem`, aplicando formateo con separador de miles al DNI (`dni.ToString("N0", _culturaEs)`, ej. `32.456.789`).
    - *Reutilización:* Selección de socios en pantallas de préstamos y operaciones regulares donde solo participan socios activos. *(Nota: Mantiene `solicitarListaSocios` como alias obsoleto por compatibilidad).*
  - **`solicitarTodosLosSocios(Optional filtro As String = "") As List(Of SocioItem)`**
    - *Firma:* `Public Function solicitarTodosLosSocios(Optional filtro As String = "") As List(Of SocioItem)`
    - *Qué hace:* Invoca `_socioDAL.buscarTodosLosSocios(filtro)` y mapea a `SocioItem` con datos completos de contacto, estado y propiedades de color/visibilidad para la UI mediante `MapearListaSociosItem`.
    - *Reutilización:* Grilla de administración general de socios (ABM) y auditoría. *(Nota: Mantiene `solicitarListaAdministrarSocios` como alias obsoleto por compatibilidad).*
  - **`solicitarSociosParaCombo(Optional filtro As String = "") As List(Of (id As Integer, nombreCompleto As String, dni As Long))`**
    - *Firma:* `Public Function solicitarSociosParaCombo(Optional filtro As String = "") As List(Of (id As Integer, nombreCompleto As String, dni As Long))`
    - *Qué hace:* Delega directamente en `_socioDAL.buscarListaSocios(filtro)`.
    - *Reutilización:* Componentes UI que requieran tuplas ligeras (ej. autocompletado en campos de texto, desplegables rápidos). *(Nota: Mantiene `solicitarListaSociosTuplas` como alias obsoleto por compatibilidad).*
  - **`cambiarEstadoSocio(idSocio As Integer, nuevoEstado As EstadoSocio) As Boolean`**
    - *Firma:* `Public Function cambiarEstadoSocio(idSocio As Integer, nuevoEstado As EstadoSocio) As Boolean`
    - *Qué hace:* Valida el identificador y actualiza el estado del socio en la base de datos a través de la DAL.
  - **`registrarSocio(nombreCompleto As String, dni As Long, email As String, telefono As String) As Socio`**
    - *Firma:* `Public Function registrarSocio(nombreCompleto As String, dni As Long, email As String, telefono As String) As Socio`
    - *Qué hace:* Valida longitud mínima de nombre, validez y unicidad del DNI, formato de email y crea el nuevo socio activo delegando a la DAL.

#### `LibroBLL.vb` (Clase)
- **Campos:**
  - `_libroDAL As New LibroDAL()`: Instancia de acceso a datos de libros.
- **Métodos:**
  - **`solicitarListaLibros(Optional filtro As String = "") As List(Of LibroItem)`**
    - *Firma:* `Public Function solicitarListaLibros(Optional filtro As String = "") As List(Of LibroItem)`
    - *Qué hace:* Obtiene los libros desde la DAL y los transforma en `LibroItem`. Aplica formateo estándar al ISBN (`FormatearIsbn`). Evalúa el stock disponible:
      - Si `stockDisponible <= 0`: `StockDisponibleTexto = "0 (Sin stock)"`, `ColorStock = "#DC2626"` (rojo), `ColorBotonSeleccionar = "#9CA3AF"`, `TieneStock = False`.
      - Si `stockDisponible = 1`: `StockDisponibleTexto = "1 ejemplar"`, `ColorStock = "#15803D"` (verde), `ColorBotonSeleccionar = "#1D3B31"`, `TieneStock = True`.
      - Si `stockDisponible > 1`: `StockDisponibleTexto = $"{stockDisponible} ejemplares"`, `ColorStock = "#15803D"`, `ColorBotonSeleccionar = "#1D3B31"`, `TieneStock = True`.
    - *Reutilización:* Pantallas de selección de libros para préstamo, vista de catálogo e inventario físico.
  - **`modificarLibro(idLibro As Integer, isbnTexto As String, titulo As String, autor As String, editorial As String, stockTotalTexto As String, Optional stockDisponibleTexto As String = "") As Libro`**
    - *Firma:* `Public Function modificarLibro(...) As Libro`
    - *Qué hace:* Valida longitud y dígitos de ISBN (13 dígitos), unicidad contra otros libros, obligatoriedad de título y autor, y que el stock total sea entero no negativo y no inferior a los préstamos activos (`stockTotal >= contarPrestamosActivos`). Recalcula automáticamente el stock disponible (`stockTotal - prestamosActivos`) para evitar desfases entre la base de datos y la realidad física de los préstamos en curso. Persiste los cambios en SQLite a través de `_libroDAL.actualizarLibro`.
  - **`registrarLibro(isbnTexto As String, titulo As String, autor As String, editorial As String, stockTotalTexto As String, Optional stockDisponibleTexto As String = "") As Libro`**
    - *Firma:* `Public Function registrarLibro(...) As Libro`
    - *Qué hace:* Valida ISBN numérico de 13 dígitos, unicidad en el catálogo, obligatoriedad de título y autor, y stock total entero $\ge 0$. Asume por regla de negocio que al dar de alta un libro nuevo, todos los ejemplares ingresan disponibles para préstamo (`stockDisponible = stockTotal`). Inserta el nuevo libro mediante `_libroDAL.guardarLibro`.
  - **`eliminarLibro(idLibro As Integer) As Boolean`**
    - *Firma:* `Public Function eliminarLibro(idLibro As Integer) As Boolean`
    - *Qué hace:* Valida la existencia del libro, comprueba que no tenga préstamos activos en curso (`contarPrestamosActivos`) y delega la baja física atómica a `_libroDAL.eliminarLibro`.
  - **`buscarLibroPorId(idLibro As Integer) As Libro`**
    - *Firma:* `Public Function buscarLibroPorId(idLibro As Integer) As Libro`
    - *Qué hace:* Valida el identificador y delega en `_libroDAL.buscarLibroPorId(idLibro)`.
  - **`existeIsbn(isbn As Long, Optional idExcluir As Integer = 0) As Boolean`**
    - *Firma:* `Public Function existeIsbn(isbn As Long, Optional idExcluir As Integer = 0) As Boolean`
    - *Qué hace:* Delega en `_libroDAL.existeIsbn(isbn, idExcluir)`.
  - **`contarPrestamosActivos(idLibro As Integer) As Integer`**
    - *Firma:* `Public Function contarPrestamosActivos(idLibro As Integer) As Integer`
    - *Qué hace:* Delega en `_libroDAL.contarPrestamosActivos(idLibro)`.
  - **`FormatearIsbn(isbn As Long) As String`** (Método Privado Compartido)
    - *Firma:* `Private Shared Function FormatearIsbn(isbn As Long) As String`
    - *Qué hace:* Convierte el código ISBN numérico de 13 dígitos en una cadena con guiones (`978-84-376-0494-7`).
  - **`solicitarLibros(Optional filtro As String = "") As List(Of Libro)`**
    - *Firma:* `Public Function solicitarLibros(Optional filtro As String = "") As List(Of Libro)`
    - *Qué hace:* Delega en `_libroDAL.buscarLibros(filtro)` retornando entidades puras de negocio.
    - *Reutilización:* Operaciones donde se requieran los objetos `Libro` completos (ej. validaciones internas o edición en formularios ABM).
  - **`solicitarLibrosDisponibles() As List(Of (id As Integer, isbn As Long, titulo As String, autor As String))`**
    - *Firma:* `Public Function solicitarLibrosDisponibles() As List(Of (id As Integer, isbn As Long, titulo As String, autor As String))`
    - *Qué hace:* Delega en `_libroDAL.buscarListaLibrosDisponibles()`.
    - *Reutilización:* Carga rápida de listas desplegables de libros con stock.

#### `PrestamoBLL.vb` (Clase)
- **Campos:**
  - `_prestamoDAL As New PrestamoDAL()`: Acceso a datos de préstamos.
  - `_socioDAL As New SocioDAL()`: Acceso a datos de socios.
  - `_libroDAL As New LibroDAL()`: Acceso a datos de libros.
  - `_culturaEs As New CultureInfo("es-AR")`: Cultura regional para formato de fechas en español.
- **Métodos:**
  - **`solicitarListaPrestamos(Optional filtro As String = "") As List(Of PrestamoItem)`**
    - *Firma:* `Public Function solicitarListaPrestamos(Optional filtro As String = "") As List(Of PrestamoItem)`
    - *Qué hace:* Consulta `_prestamoDAL.buscarPrestamos(filtro)`. Transforma cada registro a `PrestamoItem`, formateando fechas al español (`dd MMM yyyy`, ej. `06 nov 2026`). Evalúa si el préstamo venció (`p.fechaLimite.Date < Date.Today`) para asignar el color semántico de la fecha límite:
      - **Verde (`#15803D`):** Si está dentro del plazo.
      - **Rojo (`#DC2626`):** Si ya venció la fecha de devolución.
    - *Reutilización:* Vista de lista de préstamos, monitoreo de devoluciones y pantallas de control de morosidad.
  - **`registrarPrestamo(idLibro As Integer, idSocio As Integer, plazo As Integer) As Prestamo`**
    - *Firma:* `Public Function registrarPrestamo(idLibro As Integer, idSocio As Integer, plazo As Integer) As Prestamo`
    - *Qué hace:*
      1. Busca el socio por ID; si no existe, lanza `InvalidOperationException`.
      2. Busca el libro por ID; si no existe, lanza `InvalidOperationException`.
      3. Valida que el libro tenga disponibilidad (`libro.stockDisponible > 0`); de lo contrario, lanza `InvalidOperationException`.
      4. Calcula la fecha actual (`Date.Today`) y la fecha límite (`fechaActual.AddDays(plazo)`).
      5. Instancia el objeto `Prestamo` con `fechaDevolucion = Nothing` y `fueDevuelto = False`.
      6. Delega en `_prestamoDAL.guardarPrestamo(nuevoPrestamo)` para persistir en SQLite y descontar el stock.
      7. Retorna la entidad `Prestamo` completada.
    - *Reutilización:* Método central para procesar nuevos préstamos desde cualquier punto de la aplicación.
  - **`registrarDevolucion(idPrestamo As Integer) As Boolean`**
    - *Firma:* `Public Function registrarDevolucion(idPrestamo As Integer) As Boolean`
    - *Qué hace:* Delega en `_prestamoDAL.registrarDevolucion(idPrestamo)` para marcar la devolución y restituir el ejemplar al stock.
    - *Reutilización:* Acciones de devolución desde la lista de préstamos, ventanas de recepción o lectores de código de barras.

---

### 3.4. Capa de Modelos de Vista (ViewModels / DTOs de Presentación)
*Ubicación:* `Biblioteca/ViewModels/`  
Modelos específicos para enlace de datos (Data Binding) en controles XAML de WPF, desacoplando la UI del modelo relacional.

#### `SocioItem.vb`
- `IdSocio As Integer`: Identificador del socio.
- `NombreCompleto As String`: Nombre y apellido para visualización.
- `Dni As Long`: DNI numérico.
- `DniFormateado As String`: DNI formateado con puntos de miles (ej. `32.456.789`).
- `Email As String`: Dirección de correo electrónico o `"-"`.
- `Telefono As String`: Teléfono de contacto o `"-"`.
- `Estado As EstadoSocio`: Estado del socio (Activo / Inactivo).
- `EstadoTexto As String`: Texto legible (`"Habilitado"` o `"Inhabilitado"`).
- `EsActivo As Boolean`: `True` si está habilitado, `False` si está inhabilitado.
- `ColorBadgeFondo As String`: Código hexadecimal (`#E7F3EC` verde suave si activo, `#FDECE9` rojizo suave si inactivo).
- `ColorBadgeTexto As String`: Código hexadecimal (`#1D3B31` verde si activo, `#991B1B` rojo si inactivo).
- `VisibilidadBotonDesactivar As Visibility`: `Visible` si el socio está habilitado, `Collapsed` si está inhabilitado.
- `VisibilidadBotonActivar As Visibility`: `Visible` si el socio está inhabilitado, `Collapsed` si está habilitado.

#### `LibroItem.vb`
- `IdLibro As Integer`: Identificador del libro.
- `Isbn As Long`: ISBN numérico de 13 dígitos.
- `IsbnFormateado As String`: ISBN formateado con guiones (`978-84-376-0494-7`).
- `Titulo As String`: Título del libro.
- `Autor As String`: Autor.
- `Editorial As String`: Editorial.
- `StockTotal As Integer`: Cantidad física total de ejemplares.
- `StockTotalTexto As String`: Texto descriptivo (ej. `"5 ejemplares"`).
- `StockDisponible As Integer`: Stock disponible numérico.
- `StockDisponibleTexto As String`: Texto descriptivo (ej. `"5 ejemplares"`, `"1 ejemplar"`, `"0 (Sin stock)"`).
- `StockRatioTexto As String`: Texto de proporción disponible/total (ej. `"5 / 5"`).
- `ColorStock As String`: Código hexadecimal (`#15803D` en verde si hay stock, `#DC2626` en rojo si no hay).
- `ColorBadgeStock As String`: Color de fondo del badge (`#E7F3EC` verde suave si activo, `#FDECE9` rojizo suave si sin stock).
- `TieneStock As Boolean`: `True` si `StockDisponible > 0`.
- `ColorBotonSeleccionar As String`: Color para el botón/enlace de selección (`#1D3B31` o `#9CA3AF`).

#### `PrestamoItem.vb`
- `IdPrestamo As Integer`: Identificador del préstamo.
- `IdSocio As Integer`, `NombreSocio As String`, `DniSocio As Long`: Datos del socio.
- `IdLibro As Integer`, `TituloLibro As String`, `AutorLibro As String`, `IsbnLibro As Long`: Datos del libro.
- `FechaPrestamo As Date`, `FechaPrestamoFormateada As String`: Fecha de otorgamiento (`06 nov 2026`).
- `FechaLimite As Date`, `FechaLimiteFormateada As String`: Fecha de vencimiento (`20 nov 2026`).
- `FechaDevolucion As Date?`, `FechaDevolucionFormateada As String`: Fecha en que se devolvió (o vacía si sigue activo).
- `FueDevuelto As Boolean`: Estado de devolución.
- `ColorFechaLimite As String`: Color dinámico (`#15803D` verde si está en plazo, `#DC2626` rojo si venció).
- `VisibilidadBotonDevolver As Visibility` (Solo Lectura): `Visible` si `FueDevuelto = False`, `Collapsed` si ya fue devuelto.
- `VisibilidadTextoDevuelto As Visibility` (Solo Lectura): `Visible` si `FueDevuelto = True`, `Collapsed` si está pendiente.

---

### 3.5. Capa de Presentación (UI - User Interface)
*Ubicación:* `Biblioteca/`  
Vistas XAML y code-behind en VB.NET con diseño moderno en tonos verdes institucionales (`#1D3B31`, `#2D5545`, `#F7F6F2`).

#### `Application.xaml` / `Application.xaml.vb`
- **`OnStartup(e As StartupEventArgs)`:** Punto de entrada de la aplicación; invoca `DatabaseHelper.InicializarBaseDatos()` garantizando la disponibilidad de SQLite antes de levantar ventanas.

#### `MainWindow.xaml` / `MainWindow.xaml.vb`
- Ventana principal tipo Dashboard con accesos interactivos a:
  - `BtnSocios_Click`: Instancia y muestra `SociosWindow`, ocultando el menú principal.
  - `BtnLibros_Click`: Acceso al módulo de Libros (preparado para futura ventana).
  - `BtnPrestamos_Click`: Instancia y muestra `PrestamosWindow`, ocultando el menú principal.

#### `SociosWindow.xaml` / `SociosWindow.xaml.vb`
Ventana integral para la administración del estado de los socios y registro de nuevos asociados. Presenta barra de navegación lateral fija:
- **Navegación y Vistas:**
  - **`MostrarModuloListaSocios()`** (`Public Async Sub`): Despliega la grilla de administración con buscador en tiempo real, contador de socios registrados y listado detallado (Nombre, DNI, Email, Badge de Estado y Botón de Acción dinámico según estado).
  - **`MostrarModuloRegistrarSocio()`** (`Public Sub`): Despliega el formulario de alta de nuevo socio con validaciones en caliente y máscara numérica para el DNI.
- **Búsqueda Asíncrona con Debounce:**
  - **`TxtBuscarSocios_TextChanged`**: Cancela búsquedas previas y ejecuta tras 400 ms de inactividad `CargarListaSociosAsync(filtro)` buscando por nombre, DNI o email.
- **Gestión de Estados:**
  - **`BtnDesactivarSocio_Click`**: Pide confirmación y pasa el socio a estado `Inactivo` (baja lógica).
  - **`BtnActivarSocio_Click`**: Pide confirmación y reactiva el socio a estado `Activo` (habilitado).
- **Altas:**
  - **`BtnGuardarSocio_Click`**: Valida campos obligatorios y delega en `_socioBLL.registrarSocio(...)`, retornando automáticamente a la lista tras confirmar el alta exitosa.

#### `PrestamosWindow.xaml` / `PrestamosWindow.xaml.vb`
Ventana integral de administración y registro de préstamos. Implementa navegación lateral fija y flujo guiado en 3 pasos:

- **Métodos Públicos de Navegación:**
  - **`MostrarModuloListaPrestamos()`** (`Public Async Sub`): Muestra el panel `gridListaPrestamos`, resalta la opción lateral y carga los préstamos asíncronamente mediante `CargarListaPrestamosAsync()`.
  - **`MostrarPaso1Socios()`** (`Public Sub`): Muestra el panel `gridPaso1Socio` para iniciar el asistente de registro de préstamo.
- **Métodos Privados de Navegación por Pasos:**
  - **`MostrarPaso2Libros()`** (`Private Async Sub`): Muestra `gridPaso2Libro` y ejecuta `CargarLibrosAsync()`.
  - **`MostrarPaso3Plazo()`** (`Private Sub`): Muestra `gridPaso3Plazo`, actualiza la tarjeta de vista previa con los datos del socio y libro seleccionados, y marca 14 días de plazo por defecto.
- **Métodos Asíncronos de Carga con Debounce:**
  - **`CargarListaPrestamosAsync(Optional filtro As String = "") As Task`**: Ejecuta `_prestamoBLL.solicitarListaPrestamos(filtro)` en un hilo secundario (`Task.Run`) y actualiza `icListaPrestamos` y `lblContadorListaPrestamos`.
  - **`CargarSociosAsync(Optional filtro As String = "") As Task`**: Ejecuta `_socioBLL.solicitarSociosHabilitados(filtro)` con `Task.Run` y actualiza `icSocios` y `lblContadorSocios`.
  - **`CargarLibrosAsync(Optional filtro As String = "") As Task`**: Ejecuta `_libroBLL.solicitarListaLibros(filtro)` con `Task.Run` y actualiza `icLibros` y `lblContadorLibros`.
  - **`TxtBuscarPrestamos_TextChanged`**, **`TxtBuscarDni_TextChanged`**, **`TxtBuscarLibro_TextChanged`**: Manejadores de eventos de búsqueda con cancelación previa y espera de 500 ms (`CancellationTokenSource` + `Task.Delay(500)`).
- **Acciones y Transacciones:**
  - **`BtnDevolverPrestamo_Click`**: Pide confirmación al usuario y llama de forma asíncrona a `_prestamoBLL.registrarDevolucion(prestamo.IdPrestamo)`. Si es exitoso, muestra notificación y recarga la lista de préstamos y el catálogo de libros.
  - **`BtnSeleccionarSocio_Click`**: Guarda `_socioSeleccionado` y avanza al Paso 2.
  - **`BtnSeleccionarLibro_Click`**: Valida que `libro.StockDisponible > 0`, guarda `_libroSeleccionado` y avanza al Paso 3.
  - **`BtnRegistrarPrestamoFinal_Click`**: Obtiene el plazo seleccionado (7, 14 o 30 días), deshabilita el botón temporalmente y llama a `_prestamoBLL.registrarPrestamo(...)`. En caso de éxito, despliega `gridPasoExito` con el resumen del préstamo.
  - **`BtnVolverMenuPrestamos_Click`**: Limpia las variables de sesión del asistente y vuelve a mostrar el Paso 1.
  - **`ActualizarNavegacionLateral(esListaPrestamos As Boolean)`**: Alterna el estilo visual y colores de los botones laterales activos.
  - **`CerrarYVolverAlMenu()`**: Restablece y muestra la ventana `MainWindow` propietaria al cerrar la ventana de préstamos.

#### `LibrosWindow.xaml` / `LibrosWindow.xaml.vb`
Ventana integral para la administración completa del catálogo de libros (Agregar, Modificar, Eliminar). Presenta barra de navegación lateral fija:
- **Navegación y Vistas:**
  - **`MostrarModuloAgregarLibro()`** (`Public Sub`): Despliega el formulario para incorporar nuevos libros al inventario.
  - **`MostrarModuloModificarLibro()`** (`Public Async Sub`): Carga asíncronamente el catálogo mediante `CargarListaLibrosModificarAsync()` y despliega la grilla de selección y el formulario de edición.
  - **`MostrarModuloEliminarLibro()`** (`Public Async Sub`): Carga asíncronamente los libros disponibles con `CargarListaLibrosEliminarAsync()` y despliega la tabla con botones de baja definitiva.
- **Búsqueda Asíncrona con Debounce:**
  - **`TxtBuscarModificar_TextChanged`** y **`TxtBuscarEliminar_TextChanged`**: Implementan cancelación de tareas previas (`CancellationTokenSource`) y espera de 400 ms de inactividad para ejecutar consultas dinámicas por ISBN, título o autor sin saturar SQLite ni bloquear la interfaz.
- **Módulo Agregar Libro:**
  - **`BtnGuardarNuevoLibro_Click`**: Ejecuta `_libroBLL.registrarLibro(...)` validando ISBN de 13 dígitos numéricos, unicidad en catálogo, campos obligatorios de título y autor, y stock total entero $\ge 0$. Asume automáticamente que `stockDisponible = stockTotal`.
  - **`BtnLimpiarNuevoLibro_Click`**: Restablece los campos a sus valores iniciales.
- **Módulo Modificar Libro:**
  - **`BtnSeleccionarLibroModificar_Click`**: Carga el libro seleccionado en el formulario de edición y actualiza el estado visual.
  - **`BtnGuardarModificar_Click`**: Ejecuta `_libroBLL.modificarLibro(...)` con `Task.Run`. El campo `txtModificarStockDisponible` es de solo lectura (protegido); el stock disponible se recalcula automáticamente en la BLL como `stockTotal - prestamosActivos` para garantizar que la base de datos nunca se desfase con la realidad física.
  - **`BtnCancelarModificar_Click`**: Limpia el formulario y deselecciona el libro en curso.
- **Módulo Eliminar Libro:**
  - **`BtnEliminarLibro_Click`**: Solicita confirmación explícita mediante diálogo modal. Verifica mediante `_libroBLL.eliminarLibro(idLibro)` que el libro no cuente con préstamos activos (impidiendo la baja si tiene ejemplares prestados) y ejecuta una transacción atómica para remover el historial devuelto y la entidad libro. Refresca el listado tras la eliminación.

---

## 4. Guía de Reutilización para Futuras Funcionalidades

### Caso A: Implementar el Módulo de Administración de Socios (ABM)
1. **DAL (`SocioDAL.vb`):**
   - Ya se cuenta con `buscarSocios(filtro)`.
   - Se deben agregar los métodos de escritura correspondientes:
     - `guardarSocio(socio As Socio) As Integer` (con `INSERT INTO Socios (...)`).
     - `actualizarSocio(socio As Socio) As Boolean` (con `UPDATE Socios SET ...`).
     - `eliminarSocio(idSocio As Integer) As Boolean` (baja lógica: `UPDATE Socios SET estado = 1 WHERE idSocio = @id`).
2. **BLL (`SocioBLL.vb`):**
   - Reutilizar `solicitarListaSocios(filtro)` para poblar la grilla de socios.
   - Agregar validaciones de negocio: formato de DNI, unicidad de DNI, validación de email y teléfono antes de delegar a la DAL.
3. **UI (`SociosWindow.xaml`):**
   - Conectar desde `MainWindow.BtnSocios_Click`.
   - Reutilizar los estilos visuales (`PrimaryGreenButtonStyle`, bordes y sombras de tarjetas, debounce de 500 ms).

### Caso B: Módulo de Administración de Libros e Inventario (ABM) — *Completado*
1. **DAL (`LibroDAL.vb`):**
   - Métodos implementados: `buscarLibros`, `buscarLibroPorId`, `buscarListaLibrosDisponibles`, `guardarLibro`, `actualizarLibro`, `eliminarLibro`, `existeIsbn`, `contarPrestamosActivos`.
2. **BLL (`LibroBLL.vb`):**
   - Métodos implementados: `solicitarListaLibros`, `registrarLibro`, `modificarLibro`, `eliminarLibro`, `buscarLibroPorId`, `existeIsbn`, `contarPrestamosActivos`.
   - Validaciones de negocio completas: formato y unicidad de ISBN, obligatoriedad de título y autor, coherencia `stockDisponible <= stockTotal` y restricción de préstamos activos antes de modificar o eliminar.
3. **UI (`LibrosWindow.xaml` / `LibrosWindow.xaml.vb`):**
   - Conectado desde `MainWindow.BtnLibros_Click`.
   - Soporte integral de pestañas Agregar, Modificar y Eliminar con buscadores debounce de 400 ms, Data Binding dinámico y confirmaciones modales.

### Caso C: Reportes y Alertas de Morosidad o Multas
1. **BLL (`PrestamoBLL.vb`):**
   - Reutilizar `_prestamoDAL.buscarPrestamos()` y filtrar préstamos donde `fueDevuelto = False` y `fechaLimite < Date.Today`.
   - Calcular días de atraso: `(Date.Today - p.fechaLimite).Days`.
   - Calcular monto de multa en base a días de mora según política de la biblioteca.

---

## 5. Esquema de Base de Datos (SQLite)

```sql
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
```

---

## 6. Estructura de Carpetas del Proyecto

```
Biblioteca/
├── Biblioteca.slnx                  # Solución de .NET
├── GEMINI.md                        # Reglas del proyecto para asistentes IA
├── CONTEXT.md                       # Documento de contexto y arquitectura (este archivo)
└── Biblioteca/                      # Proyecto principal WPF (.NET 10 VB.NET)
    ├── Biblioteca.vbproj            # Archivo de configuración del proyecto
    ├── Application.xaml             # Definición de aplicación WPF
    ├── Application.xaml.vb          # Inicialización de DB en OnStartup
    ├── AssemblyInfo.vb              # Metadatos del ensamblado
    ├── DatabaseHelper.vb            # Helper de base de datos SQLite y DDL
    ├── biblioteca.db                # Base de datos SQLite local
    │
    ├── BE / Models/                 # Capa BE (Entidades de Negocio)
    │   ├── Socio.vb                 # Entidad Socio y Enum EstadoSocio
    │   ├── Libro.vb                 # Entidad Libro
    │   └── Prestamo.vb              # Entidad Prestamo
    │
    ├── DAL/                         # Capa DAL (Acceso a Datos)
    │   ├── SocioDAL.vb              # Acceso a datos de Socios
    │   ├── LibroDAL.vb              # Acceso a datos de Libros
    │   └── PrestamoDAL.vb           # Acceso a datos de Préstamos
    │
    ├── BLL/                         # Capa BLL (Lógica de Negocio)
    │   ├── SocioBLL.vb              # Reglas de negocio de Socios
    │   ├── LibroBLL.vb              # Reglas de negocio de Libros
    │   └── PrestamoBLL.vb           # Reglas de negocio de Préstamos
    │
    ├── ViewModels/                  # Modelos de Vista / DTOs para la UI
    │   ├── SocioItem.vb             # Item formateado para listados de Socios
    │   ├── LibroItem.vb             # Item formateado para listados de Libros
    │   └── PrestamoItem.vb          # Item formateado para listados de Préstamos
    │
    └── UI / Vistas/                 # Capa UI (Presentación)
        ├── MainWindow.xaml          # Menú principal y navegación
        ├── MainWindow.xaml.vb       # Code-behind del menú principal
        ├── SociosWindow.xaml        # Módulo de administración y alta de socios
        ├── SociosWindow.xaml.vb     # Code-behind de socios con búsqueda debounce y gestión de estados
        ├── LibrosWindow.xaml        # Módulo de administración de libros (Agregar, Modificar, Eliminar)
        ├── LibrosWindow.xaml.vb     # Code-behind de administración de libros y navegación lateral
        ├── PrestamosWindow.xaml     # Gestión y registro de préstamos
        └── PrestamosWindow.xaml.vb  # Code-behind con búsqueda asíncrona y debounce
```

---

## 7. Convenciones y Reglas de Desarrollo

1. **Lenguaje Exclusivo:** Todo el código nuevo o modificado debe redactarse exclusivamente en **Visual Basic .NET (`.vb`)**.
2. **Respeto a la Arquitectura en Capas:**
   - La capa **UI** sólo debe comunicarse con la capa **BLL** (nunca consultar la base de datos o la DAL directamente).
   - La capa **BLL** coordina validaciones, transformaciones a ViewModels y consulta la capa **DAL**.
   - La capa **DAL** es la única responsable de sentencias SQL y conexiones con SQLite.
   - Las entidades **BE** fluyen a través de todas las capas como portadoras de datos.
3. **Seguridad en Consultas:** Usar siempre parámetros SQL (`SqliteParameter` / `AddWithValue`) para mitigar riesgos de inyección SQL.
4. **Experiencia de Usuario (UX) Asíncrona:** Operaciones de consulta o persistencia pesadas deben ejecutarse de forma asíncrona (`Async`/`Await`), implementando debouncing en cuadros de búsqueda.
