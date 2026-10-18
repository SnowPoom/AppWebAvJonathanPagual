<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
' Agrega una canción a una playlist propia (JSON)
Response.ContentType = "application/json"

Dim idPl, idSong, owns, exists, songExists, plTitle

If Not IsLoggedIn() Then
    Response.Status = "401 Unauthorized"
    Response.Write "{""ok"":false,""message"":""Sesión no iniciada.""}"
    Response.End
End If

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Status = "405 Method Not Allowed"
    Response.Write "{""ok"":false,""message"":""Método no permitido.""}"
    Response.End
End If

If Not CsrfValid() Then
    Response.Status = "403 Forbidden"
    Response.Write "{""ok"":false,""message"":""Token de seguridad invalido.""}"
    Response.End
End If

idPl = ToInt(FM("playlist"), 0)
idSong = ToInt(FM("song"), 0)

If idPl <= 0 Or idSong <= 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""message"":""Datos incompletos.""}"
    Response.End
End If

owns = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ? AND ID_USER = ?", Array(idPl, CurrentUserId())), 0)
If owns = 0 Then
    Response.Status = "403 Forbidden"
    Response.Write "{""ok"":false,""message"":""La playlist no pertenece al usuario.""}"
    Response.End
End If

songExists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong)), 0)
If songExists = 0 Then
    Response.Status = "404 Not Found"
    Response.Write "{""ok"":false,""message"":""La canción no existe.""}"
    Response.End
End If

exists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(idPl, idSong)), 0)
If exists > 0 Then
    Response.Write "{""ok"":true,""message"":""La canción ya estaba en esa playlist.""}"
    Response.End
End If

Call DbExecute("INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", Array(idPl, idSong))
plTitle = DbScalar("SELECT TITLE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(idPl)) & ""

Response.Write "{""ok"":true,""message"":" & JsonStr("Canción agregada a '" & plTitle & "'.") & "}"
DbClose()
%>
