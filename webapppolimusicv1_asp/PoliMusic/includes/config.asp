<%
' =====================================================================
'  PoliMusic - Configuracion central del sitio
'  ASP Clasico (VBScript) - sin dependencias externas
' =====================================================================
Response.Buffer   = True
Response.CodePage = 65001
Response.Charset  = "utf-8"
Session.CodePage  = 65001
Session.Timeout   = 60

Const APP_NAME    = "PoliMusic"
Const APP_TAGLINE = "Tu universo sonoro"
Const APP_VERSION = "1.0.0"

' ---------------------------------------------------------------------
'  Cadena de conexion.
'  Los datos sensibles viven en includes/config.local.asp, que NO se
'  versiona (ver .gitignore). Se crea copiando config.example.asp y se
'  ajusta con los datos del servidor. Aqui solo se declara la variable.
' ---------------------------------------------------------------------
Dim DB_CONN_STRING
%>
<!--#include file="config.local.asp"-->
<%
' --- Constantes ADO (Classic ASP no referencia la type library) ---
Const adCmdText     = 1
Const adOpenStatic  = 3
Const adLockReadOnly = 1
Const adUseClient   = 3
Const adParamInput  = 1
Const adInteger     = 3
Const adVarChar     = 200
Const adVarWChar    = 202
Const adDBTimeStamp = 135
Const adDate        = 7
Const adDouble      = 5
Const adBoolean     = 11
Const adExecuteNoRecords = 128

' --- Roles (TBL_CATALOG -> 1/0 = ADMIN, 1/1 = USER) ---
Const ROLE_USER  = 0
Const ROLE_ADMIN = 1

' --- Carpetas fisicas de recursos: coinciden con los valores de la BD ---
'     SONG_PATH        -> ~/songFiles/...
'     ALBUM_IMAGE_PATH -> ~/albumImages/...
'     ARTIST_IMAGE     -> ~/ArtistImages/...
'     PHOTO (playlist) -> ~/playlistImages/...
'     USER_PHOTO       -> ~/userImages/...
Const DIR_SONGS     = "songFiles"
Const DIR_ALBUMS    = "albumImages"
Const DIR_ARTISTS   = "artistImages"
Const DIR_PLAYLISTS = "playlistImages"
Const DIR_USERS     = "userImages"

' --- Paginacion / limites de interfaz ---
Const PAGE_SIZE_DEFAULT = 8

Dim gConn, gAppBase, gFso
Set gConn = Nothing
Set gFso  = Nothing
gAppBase  = ""
%>
