# PoliMusic

Aplicación web de música desarrollada **exclusivamente con ASP Clásico (VBScript)**, HTML, CSS y JavaScript sobre **SQL Server**. No utiliza .NET, ni frameworks, ni dependencias externas de frontend.

Identidad visual propia: paleta de azules, celestes y verdes menta sobre una interfaz **Glassmorphism** (paneles translúcidos, desenfoque de fondo, degradados aurora y microanimaciones).

---

## 1. Requisitos

| Componente | Versión / nota |
|---|---|
| IIS | 10 (Windows 10/11 o Windows Server) con la característica **ASP** habilitada |
| SQL Server | 2016 o superior (probado en SQL Server 2025 Express) |
| Navegador | Chrome, Edge o Firefox actuales (usa `backdrop-filter`) |
| Base de datos | Script `BDD_PoliMusic.sql` |

La configuración de la base de datos **no está en el repositorio**: vive en `includes/config.local.asp`, un archivo ignorado por Git. Antes de abrir la aplicación hay que crearlo (ver el punto siguiente).

---

## 2. Configuración de la conexión (obligatorio)

1. En `PoliMusic/includes/`, copie la plantilla:

   ```cmd
   copy config.example.asp config.local.asp
   ```

   ```powershell
   Copy-Item config.example.asp config.local.asp
   ```

2. Edite `config.local.asp` con los datos de su servidor:

   | Variable | Ejemplo | Descripción |
   |---|---|---|
   | `DB_SERVER` | `localhost` | Servidor o instancia (`.\SQLEXPRESS`, `192.168.1.10`…) |
   | `DB_PORT` | `14333` | Puerto TCP; use `""` para el puerto por defecto |
   | `DB_NAME` | `BDD_PoliMusic` | Base creada por `BDD_PoliMusic.sql` |
   | `DB_USER` | `sa` | Usuario SQL (también sirve `usr_polimusic`) |
   | `DB_PASSWORD` | `••••••••` | Contraseña de ese usuario |

3. La cadena final se arma sola con esos valores y queda disponible como `DB_CONN_STRING` para todo el sitio:

   ```
   Provider=MSOLEDBSQL;Data Source=<DB_SERVER>,<DB_PORT>;Initial Catalog=<DB_NAME>;
   User ID=<DB_USER>;Password=<DB_PASSWORD>;Encrypt=True;TrustServerCertificate=True;
   Persist Security Info=True;Pooling=False;MultipleActiveResultSets=False;
   Application Name="PoliMusic"
   ```

> **No versione `config.local.asp`**: ya está en `.gitignore`. Si lo pierde, vuelva a copiar la plantilla; nunca lo suba al repositorio.
>
> Si el archivo no existe, IIS devolverá el error de compilación **ASP 0126** indicando que falta el archivo de inclusión.

Se usa el proveedor **MSOLEDBSQL** porque la cadena requiere `Encrypt=True`, opción que el proveedor antiguo `SQLOLEDB` no admite. El script de la base también crea el login `usr_polimusic` (db_owner) como alternativa a `sa`.

---

## 3. Estructura del proyecto

```
PoliMusic/
├── default.asp                  Punto de entrada (redirige según la sesión)
├── login.asp  logout.asp        Autenticación con sesiones ASP
├── home.asp                     Pantalla principal con mensaje de bienvenida
├── songs.asp  artists.asp  albums.asp  search.asp
├── artist.asp  album.asp        Vistas de detalle
├── playlists.asp  playlist.asp  Playlists del usuario (CRUD)
├── profile.asp                  Datos personales y cambio de contraseña
├── admin/
│   ├── default.asp              Panel con métricas y estado de archivos
│   ├── users.asp   user_form.asp
│   ├── songs.asp   song_form.asp
│   ├── artists.asp artist_form.asp
│   └── albums.asp  album_form.asp
├── api/
│   ├── play.asp                 Incrementa TBL_SONG.PLAYS (JSON)
│   └── playlist_add.asp         Agrega canción a una playlist propia (JSON)
├── includes/
│   ├── config.asp               Constantes, rutas e inclusión de config.local.asp
│   ├── config.example.asp       Plantilla de conexión (sí se versiona)
│   ├── config.local.asp         Conexión real (NO se versiona; se crea al instalar)
│   ├── functions.asp            Helpers (rutas, HTML, JSON, ADO, iconos SVG)
│   ├── auth.asp                 Sesión, roles y verificación de credenciales
│   ├── ui.asp                   Componentes (fila de canción, tarjetas)
│   ├── layout_top.asp           Cabecera, sidebar y topbar
│   └── layout_bottom.asp        Reproductor global y scripts
├── assets/
│   ├── css/polimusic.css        Sistema de diseño completo
│   ├── css/fonts.css            Tipografías locales (Sora + Manrope)
│   ├── fonts/*.woff2            16 archivos (340 KB, sin CDN)
│   ├── js/polimusic.js          Reproductor y comportamiento de la interfaz
│   └── img/favicon.svg
├── songFiles/       bensound-adventure.mp3
├── albumImages/     (vacía: se generan portadas de respaldo)
├── artistImages/    (vacía: se generan avatares de respaldo)
├── playlistImages/  default.jpg + playlist1..5.png
├── userImages/      default.jpg + root.png
└── web.config       Documento predeterminado, MIME y cabeceras
```

---

## 4. Análisis de la base de datos

El esquema se respetó **sin agregar ni modificar tablas, campos ni relaciones**.

| Tabla | Función | Campos relevantes |
|---|---|---|
| `TBL_CATALOG` | Dominios (tipos) | `COD_CATALOG_PARENT`, `COD_CATALOG`, `VALUE` |
| `TBL_USER` | Cuentas | `USERNAME`, `PASSWORD`, `EMAIL`, `BIRTHDAY`, `USER_PHOTO`, `USER_TYPE` |
| `TBL_SONG` | Canciones | `SONG_NAME`, `SONG_PATH`, `PLAYS` |
| `TBL_ARTIST` | Artistas | `ARTIST_NAME`, `ARTIST_LASTNAME`, `ARTIST_IMAGE` |
| `TBL_ALBUM` | Álbumes | `TITLE_ALBUM`, `RELEASE_YEAR`, `ALBUM_IMAGE_PATH` |
| `TBL_TRACKLIST` | Álbum ↔ Canción | PK compuesta (`ID_ALBUM`, `ID_SONG`) |
| `TBL_PLAYER` | Canción ↔ Artista | PK compuesta (`ID_SONG`, `ID_ARTIST`) |
| `TBL_PLAYLIST` | Playlists | `ID_USER`, `TITLE`, `CREATION_DATE`, `TYPE`, `PHOTO` |
| `TBL_PLAYLIST_SONG` | Playlist ↔ Canción | PK compuesta (`ID_PLAYLIST`, `ID_SONG`) |

**Dominios (TBL_CATALOG):**

- Grupo `1` → `0 = USER`, `1 = ADMIN` (se usa en `TBL_USER.USER_TYPE`).
- Grupo `2` → `0 = Public`, `1 = Private` (se usa en `TBL_PLAYLIST.TYPE`).

**Datos existentes utilizados:** 3 usuarios, 5 canciones, 5 artistas, 5 álbumes, 5 playlists, 13 pistas de playlist, 14 pistas de álbum y 10 relaciones canción-artista.

**Contraseñas:** la semilla contiene dos en texto plano (`12345678`) y una en SHA-1. El login acepta ambas formas y **todo valor nuevo se guarda como SHA-1** calculado por SQL Server (`HASHBYTES`), sin depender de componentes COM.

**Relaciones al eliminar:** el panel administrativo borra primero las filas dependientes (`TBL_PLAYER`, `TBL_TRACKLIST`, `TBL_PLAYLIST_SONG`, `TBL_PLAYLIST`) dentro de una transacción, para no violar las claves foráneas.

---

## 5. Cuentas de prueba

| Usuario | Contraseña | Rol |
|---|---|---|
| `userAdmin` | `12345678` | Administrador |
| `userNormal` | `12345678` | Usuario |
| `Jonathan.Pagual` | `12345678` | Usuario |

---

## 6. Funcionalidades

### Visitante
- Login con validación de credenciales, bloqueo temporal tras 5 intentos fallidos y mensajes de error claros.

### Usuario normal
- Inicio con **mensaje de bienvenida personalizado** usando `Session("Usuario")`, saludo por franja horaria y métricas del catálogo.
- Explorar y buscar canciones (por título, artista o álbum), artistas y álbumes, con filtros y ordenamientos.
- Vistas de detalle de artista y álbum con sus canciones y álbumes relacionados.
- Buscador global en la barra superior.
- **Playlists propias**: crear, editar (título, visibilidad y portada), eliminar, agregar y quitar canciones.
- Ver playlists públicas de otros usuarios; las privadas solo son visibles para su dueño.
- Perfil: actualizar nombre, correo, fecha de nacimiento, foto y contraseña.
- Reproductor con cola, aleatorio, repetición, volumen y progreso persistente entre páginas.

### Administrador
- Todo lo anterior, más el módulo de administración:
  - **Usuarios**: listar, buscar, crear, editar (rol incluido) y eliminar. Protecciones: no puede eliminarse a sí mismo ni dejar el sistema sin administradores; al borrar un usuario se eliminan sus playlists.
  - **Canciones**: CRUD completo, relación con artistas (`TBL_PLAYER`) y álbumes (`TBL_TRACKLIST`) mediante selección múltiple, y verificación en disco del archivo de audio.
  - **Artistas**: CRUD con imagen y conteo de canciones/álbumes.
  - **Álbumes**: CRUD con portada, año y lista de canciones.
  - **Panel**: métricas generales, canciones más escuchadas, usuarios recientes y **estado real de los archivos de audio**.

### Reproductor
Play/pause, anterior, siguiente, aleatorio, repetir (todo / una), barra de progreso con arrastre y teclado, volumen con silencio, cola visible, portada + título + artista, ecualizador animado, atajos (`Espacio`, `←/→` con Shift, `/` para buscar) y **Media Session API**. El contador `PLAYS` se incrementa en cada reproducción mediante `api/play.asp`.

Las pistas cuyo archivo no existe se marcan como **“Sin audio”**, no se reproducen y el botón *siguiente* las salta automáticamente.

---

## 7. Seguridad implementada

- Consultas **parametrizadas** con `ADODB.Command` en toda la aplicación (previene inyección SQL).
- Escape de salida HTML (`HtmlEncode`) y de cadenas JSON.
- **Token CSRF** por sesión en todos los formularios y llamadas `fetch`.
- Control de acceso por rol en cada página (`RequireLogin`, `RequireAdmin`) y validación de propiedad en playlists.
- Validación de servidor de todos los campos (longitudes, correo, fechas, IDs).
- Redirección `next` restringida a rutas internas (evita *open redirect*).
- Contraseñas nuevas con SHA-1; bloqueo temporal de login tras intentos fallidos.
- Cabeceras `X-Content-Type-Options`, `Referrer-Policy`, `X-Frame-Options` y exclusión de `X-Powered-By`.

---

## 8. Recursos multimedia

La carpeta `SongFilesAndImages_resources` contenía:

- `SongFiles/bensound-adventure.mp3` → copiado a `songFiles/`.
- `UserImages/default.jpg`, `root.png` → copiados a `userImages/`.
- `PlaylistImages/default.jpg`, `playlist1..5.png` → copiados a `playlistImages/`.

**No existían** los otros cuatro MP3 (`dreams`, `energy`, `scifi`, `tomorrow`) ni las carpetas `albumImages` / `ArtistImages` referenciadas por la base. Para no inventar contenido:

- Las canciones sin archivo muestran la etiqueta *Sin audio* y el reproductor las omite.
- Los álbumes y artistas sin imagen usan una **portada generada** con degradado de la identidad visual e iniciales del nombre.

Para completar el catálogo basta copiar los MP3 en `songFiles/` (mismos nombres que `TBL_SONG.SONG_PATH`) y las imágenes en `albumImages/`, `artistImages/`, sin tocar la base de datos.

---

## 9. Despliegue

Ver **[DEPLOY.md](DEPLOY.md)** para el paso a paso en IIS y la solución de problemas.

---

## 10. Pruebas realizadas

Verificado sobre IIS 10 con SQL Server 2025:

- Login de usuario normal y administrador, credenciales incorrectas, bloqueo por intentos y logout.
- Acceso denegado de un usuario normal a `/admin/*` y a playlists privadas ajenas.
- CRUD completo de usuarios, canciones, artistas y álbumes, incluidas las relaciones y el borrado en cascada.
- CRUD de playlists y de sus canciones (alta, baja y duplicados).
- Cambio de contraseña (con validación de la actual) y nuevo inicio de sesión con el hash SHA-1.
- API de reproducciones y de agregado a playlist, incluidos los rechazos por CSRF, propiedad y datos inválidos.
- Renderizado correcto en 1440 px, 1024 px y 390 px, sin errores de consola.
- Publicación tanto en la raíz del sitio (`/`) como en carpeta (`/PoliMusic/`), con rutas y recursos resueltos en ambos casos.
