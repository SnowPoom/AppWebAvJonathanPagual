<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<%
' Registra una reproducción y devuelve el contador actualizado (JSON)
Response.ContentType = "application/json"

Dim idSong, plays
idSong = ToInt(Request.QueryString("id"), 0)

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

If idSong <= 0 Then
    Response.Status = "400 Bad Request"
    Response.Write "{""ok"":false,""message"":""Canción no válida.""}"
    Response.End
End If

Call DbExecute("UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array(idSong))
plays = ToInt(DbScalar("SELECT ISNULL(PLAYS, 0) FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong)), 0)

Response.Write "{""ok"":true,""plays"":" & plays & "}"
DbClose()
%>
