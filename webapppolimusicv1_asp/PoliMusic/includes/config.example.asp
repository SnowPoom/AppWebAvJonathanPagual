<%
' =====================================================================
'  PoliMusic - PLANTILLA de configuracion local
'
'  Copie este archivo como  config.local.asp  en esta misma carpeta y
'  ajuste los valores.  config.local.asp NO se versiona porque contiene
'  credenciales (ver .gitignore).
'
'      CMD:         copy config.example.asp config.local.asp
'      PowerShell:  Copy-Item config.example.asp config.local.asp
'
'  Si config.local.asp no existe, las paginas devuelven el error de
'  compilacion ASP 0126 (archivo de inclusion no encontrado).
'
'  NOTA: no declare DB_CONN_STRING con Dim; ya se declara en config.asp.
' =====================================================================

' --- Datos de conexion a SQL Server ---
Dim DB_SERVER, DB_PORT, DB_NAME, DB_USER, DB_PASSWORD, DB_DATASOURCE

DB_SERVER   = "localhost"           ' servidor o instancia: "localhost", ".\SQLEXPRESS", "192.168.1.10"
DB_PORT     = "1433"                ' puerto TCP; use "" para el puerto por defecto
DB_NAME     = "BDD_PoliMusic"       ' nombre de la base creada por BDD_PoliMusic.sql
DB_USER     = "sa"                  ' usuario SQL; tambien sirve "usr_polimusic"
DB_PASSWORD = "CAMBIA_ESTA_CLAVE"   ' contrasena del usuario anterior

' --- Cadena final (normalmente no hace falta modificarla) ---
If DB_PORT = "" Then
    DB_DATASOURCE = DB_SERVER
Else
    DB_DATASOURCE = DB_SERVER & "," & DB_PORT
End If

DB_CONN_STRING = "Provider=MSOLEDBSQL;Data Source=" & DB_DATASOURCE & _
                 ";Initial Catalog=" & DB_NAME & ";User ID=" & DB_USER & ";Password=" & DB_PASSWORD & _
                 ";Encrypt=True;TrustServerCertificate=True;Persist Security Info=True;Pooling=False;" & _
                 "MultipleActiveResultSets=False;Application Name=""PoliMusic"""
%>
