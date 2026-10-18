# Despliegue de PoliMusic en IIS

Guía para publicar la aplicación en **`C:\inetpub\wwwroot\PoliMusic`** y ejecutarla con ASP Clásico + SQL Server.

---

## 0. Antes de empezar

1. La base de datos debe existir. Ejecute `BDD_PoliMusic.sql` en SQL Server (crea `BDD_PoliMusic`, las 9 tablas y los datos de ejemplo).
2. Cree el archivo de conexión local (no viene en el repositorio):

   ```cmd
   cd PoliMusic\includes
   copy config.example.asp config.local.asp
   ```

   Luego edite `config.local.asp` con `DB_SERVER`, `DB_PORT`, `DB_NAME`, `DB_USER` y `DB_PASSWORD`. **No modifique `config.asp`.**
3. El sitio se publica en una **carpeta**, por lo que la URL final es `http://localhost/PoliMusic/`. La aplicación calcula su ruta virtual automáticamente, así que también funciona publicada en la raíz del sitio.

---

## 1. Despliegue automático (recomendado)

Desde una consola **PowerShell ejecutada como Administrador**:

```powershell
cd D:\webAv\AppWebAvJonathanPagual\webapppolimusicv1_asp
powershell -ExecutionPolicy Bypass -File .\deploy_iis.ps1
```

El script realiza todo lo siguiente:

1. Habilita la característica **IIS-ASP** si no está activa.
2. Copia el contenido de `PoliMusic\` a `C:\inetpub\wwwroot\PoliMusic` (robocopy `/MIR`).
3. Concede lectura a `IIS_IUSRS` y control total al usuario actual (para futuras actualizaciones).
4. Configura ASP a nivel de servidor: sesión de 60 minutos, **rutas primarias habilitadas** y errores detallados.
5. Se asegura de que el *app pool* y el sitio estén iniciados.
6. Verifica por HTTP que `login.asp` y la hoja de estilos respondan.

Al finalizar, abra `http://localhost/PoliMusic/`.

> Para volver a publicar tras un cambio basta con ejecutar de nuevo el script (o copiar los archivos modificados).

---

## 2. Despliegue manual

### 2.1 Habilitar ASP Clásico

Panel de control → *Programas* → *Activar o desactivar características de Windows* → **Servicios de Internet Information Server → Servicios World Wide Web → Características de desarrollo de aplicaciones → ASP**.

O por línea de comandos (como administrador):

```cmd
dism /online /enable-feature /featurename:IIS-ASP /all /norestart
```

### 2.2 Copiar los archivos

Copie **el contenido** de la carpeta `PoliMusic` (no la carpeta) en:

```
C:\inetpub\wwwroot\PoliMusic\
```

Debe quedar `C:\inetpub\wwwroot\PoliMusic\login.asp`, `...\includes\config.asp`, `...\assets\...`, etc.

### 2.3 Crear la aplicación en IIS

**Opción A — como carpeta del sitio por defecto (más simple):** no requiere nada adicional, ya que `C:\inetpub\wwwroot` es la raíz de *Default Web Site*.

**Opción B — como aplicación IIS:** Administrador de IIS → *Default Web Site* → clic derecho en la carpeta `PoliMusic` → **Convertir en aplicación** → seleccionar un *app pool* con **.NET CLR = Sin código administrado** → Aceptar.

### 2.4 App pool

- Versión de .NET CLR: **Sin código administrado**.
- Modo de canalización: **Integrado**.
- Identidad: `ApplicationPoolIdentity` (predeterminada) es suficiente.

### 2.5 Permisos NTFS

El grupo **`IIS_IUSRS`** necesita **Lectura y ejecución** sobre `C:\inetpub\wwwroot\PoliMusic`. Normalmente ya se hereda de `wwwroot`; si no:

```cmd
icacls "C:\inetpub\wwwroot\PoliMusic" /grant "IIS_IUSRS:(OI)(CI)(RX)" /T
```

La aplicación **no escribe archivos** (las rutas de audio e imágenes se guardan en la base de datos), por lo que no hacen falta permisos de escritura.

### 2.6 Rutas primarias (obligatorio)

Las páginas de `admin\` incluyen los archivos de `..\includes\`. IIS bloquea `..` por defecto y devuelve **ASP 0131**. Habilítelo así:

Administrador de IIS → seleccione el servidor → *Editor de configuración* → sección **`system.webServer/asp`** → `enableParentPaths = True`.

O por línea de comandos (como administrador):

```cmd
%windir%\system32\inetsrv\appcmd.exe set config /section:asp /enableParentPaths:"true" /commit:apphost
```

### 2.7 Sesión y errores

```cmd
%windir%\system32\inetsrv\appcmd.exe set config /section:asp /session.timeout:"00:01:00" /commit:apphost
%windir%\system32\inetsrv\appcmd.exe set config /section:asp /scriptErrorSentToBrowser:"true" /commit:apphost
```

- `session.timeout` alinea la sesión de ASP con los 60 minutos configurados en `config.asp`.
- `scriptErrorSentToBrowser` muestra el detalle de los errores ASP. **En producción** cámbielo a `false` y revise el registro de eventos.

### 2.8 Documento predeterminado, MIME y cabeceras

Ya están resueltos por `web.config` (incluido en la carpeta):

- Documento predeterminado: `default.asp`.
- Tipos MIME para `.mp3`, `.m4a`, `.ogg`, `.wav`, `.webp`, `.svg`, `.woff2`.
- Cabeceras `X-Content-Type-Options`, `Referrer-Policy`, `X-Frame-Options` y eliminación de `X-Powered-By`.

Si IIS muestra **500.19** al abrir el sitio, comente el bloque `<staticContent>` de `web.config` y agregue los MIME desde el Administrador de IIS (*Tipos MIME*), ya que algunos servidores bloquean esa sección.

---

## 3. Verificación

1. `http://localhost/PoliMusic/` debe redirigir a `login.asp`.
2. Inicie sesión con `userAdmin` / `12345678`.
3. Compruebe en el panel que las 5 canciones, 5 artistas y 5 álbumes aparecen.
4. Reproduzca **Adventure**: es la única pista con archivo MP3 real en el paquete.
5. Entre a `admin/` y cree/edite un registro para confirmar escritura en la base.

---

## 4. Solución de problemas

| Síntoma | Causa probable | Solución |
|---|---|---|
| **ASP 0126 – No se puede encontrar el archivo incluido** | Falta `includes/config.local.asp` | Paso 0.2: copie `config.example.asp` como `config.local.asp` |
| La página se descarga como texto o da 404.3 | ASP no está habilitado | Instale la característica **IIS-ASP** (paso 2.1) |
| **ASP 0131 – Ruta de directorio primario no autorizada** | Rutas primarias deshabilitadas | Paso 2.6 |
| **500.19** al abrir cualquier página | `web.config` con una sección bloqueada | Comente `<staticContent>` y configure los MIME desde IIS |
| **Error '800a0e7a' / no se encuentra el proveedor** | Falta MSOLEDBSQL | Instale *Microsoft OLE DB Driver for SQL Server* o cambie el proveedor en `config.asp` |
| Error de inicio de sesión de SQL Server | Credenciales o puerto incorrectos | Ajuste `DB_CONN_STRING`; el servidor de ejemplo escucha en `localhost,14333` |
| El audio no suena | El MP3 no existe en `songFiles` | Copie el archivo con el nombre exacto de `TBL_SONG.SONG_PATH` |
| Se pierde la sesión rápido | Timeout de ASP | Paso 2.7 (`session.timeout`) |
| Los iconos se ven como cuadros | Falta la carpeta `assets/fonts` | Copie el proyecto completo |

---

## 5. Notas para producción

- Cambie la contraseña de `sa` en `includes/config.asp` por un usuario de base de datos con permisos mínimos (por ejemplo `usr_polimusic`, ya creado por el script).
- Desactive `scriptErrorSentToBrowser` y `httpErrors errorMode="DetailedLocalOnly"`.
- Publique sobre HTTPS y active *HSTS* si el sitio es público.
- El hash de contraseñas es **SHA-1** por compatibilidad con los datos existentes; para un entorno real se recomienda migrar a un algoritmo con *salt* (por ejemplo PBKDF2 en la capa de aplicación).
- Programe copias de seguridad de `BDD_PoliMusic` y de las carpetas `songFiles`, `albumImages`, `artistImages`, `playlistImages` y `userImages`.
