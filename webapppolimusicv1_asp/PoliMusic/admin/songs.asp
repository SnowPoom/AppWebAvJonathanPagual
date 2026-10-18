<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Canciones"
pageSubtitle = "Catálogo administrable de PoliMusic"
navActive = "admin_songs"
RequireAdmin()

Dim q, rs, action, idSong, targetName, conn, errMsg, sql, prmArr
action = FM("action")
errMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not CsrfValid() Then
        Call Flash("La sesión expiró. Vuelve a intentarlo.", "error")
        GoToPage AppBase() & "admin/songs.asp"
    ElseIf action = "delete" Then
        idSong = ToInt(FM("id"), 0)
        targetName = DbScalar("SELECT SONG_NAME FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong)) & ""
        If idSong <= 0 Or targetName = "" Then
            Call Flash("La canción indicada no existe.", "error")
        Else
            Set conn = Db()
            On Error Resume Next
            conn.BeginTrans
            Call DbExecute("DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(idSong))
            Call DbExecute("DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(idSong))
            Call DbExecute("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?", Array(idSong))
            Call DbExecute("DELETE FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong))
            If Err.Number <> 0 Then
                conn.RollbackTrans
                errMsg = "No se pudo eliminar la canción: " & Err.Description
                Err.Clear
            Else
                conn.CommitTrans
                Call Flash("Canción '" & targetName & "' eliminada del catálogo.", "success")
            End If
            On Error GoTo 0
        End If
        GoToPage AppBase() & "admin/songs.asp"
    End If
End If

q = QS("q")
sql = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS,0) AS PLAYS, " & _
      "ISNULL((SELECT STRING_AGG(LTRIM(RTRIM(a.ARTIST_NAME + ' ' + ISNULL(a.ARTIST_LASTNAME,''))), ', ') " & _
      "  FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG), '') AS ARTISTS, " & _
      "ISNULL((SELECT STRING_AGG(al.TITLE_ALBUM, ', ') FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al ON al.ID_ALBUM = t.ID_ALBUM " & _
      "  WHERE t.ID_SONG = s.ID_SONG), '') AS ALBUMS " & _
      "FROM TBL_SONG s "
If q <> "" Then
    sql = sql & "WHERE s.SONG_NAME LIKE ? ORDER BY s.SONG_NAME"
    Set rs = DbQuery(sql, Array(LikeParam(q)))
Else
    sql = sql & "ORDER BY s.ID_SONG"
    Set rs = DbQuery(sql, Empty)
End If
%>
<!--#include file="../includes/layout_top.asp"-->

<% If errMsg <> "" Then %>
<div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
<% End If %>

<div class="admin-head">
    <div>
        <h2><%=rs.RecordCount%>&#32;<%=IIf(rs.RecordCount = 1, "canción", "canciones")%></h2>
        <p>Cada registro se relaciona con artistas (TBL_PLAYER) y álbumes (TBL_TRACKLIST).</p>
    </div>
    <div class="row">
        <form method="get" action="<%=AppBase()%>admin/songs.asp" class="row" style="gap:6px">
            <input class="input" type="search" name="q" value="<%=HtmlAttr(q)%>" placeholder="Buscar canción" style="height:40px;min-width:190px">
            <button class="btn btn--ghost btn--sm" type="submit"><%=Icon("search", "")%></button>
        </form>
        <a class="btn btn--primary" href="<%=AppBase()%>admin/song_form.asp"><%=Icon("plus", "")%> Nueva canción</a>
    </div>
</div>

<div class="table-wrap">
    <table class="data">
        <thead>
            <tr>
                <th>ID</th>
                <th>Canción</th>
                <th>Artistas</th>
                <th>Álbumes</th>
                <th class="num">Reproducciones</th>
                <th>Archivo</th>
                <th style="text-align:right">Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% Do Until rs.EOF %>
            <tr>
                <td class="num muted"><%=rs("ID_SONG")%></td>
                <td><strong><%=HtmlEncode(rs("SONG_NAME") & "")%></strong>
                    <div class="small muted mono"><%=HtmlEncode(rs("SONG_PATH") & "")%></div>
                </td>
                <td class="small"><%=HtmlEncode(TextOr(rs("ARTISTS") & "", "-"))%></td>
                <td class="small"><%=HtmlEncode(TextOr(rs("ALBUMS") & "", "-"))%></td>
                <td class="num"><%=ToInt(rs("PLAYS"), 0)%></td>
                <td>
                    <% If AssetExists(rs("SONG_PATH") & "") Then %>
                    <span class="chip chip--admin">OK</span>
                    <% Else %>
                    <span class="chip chip--muted">Falta</span>
                    <% End If %>
                </td>
                <td>
                    <div class="cell-actions">
                        <a class="icon-btn" href="<%=AppBase()%>admin/song_form.asp?id=<%=rs("ID_SONG")%>" title="Editar"><%=Icon("edit", "")%></a>
                        <form method="post" action="<%=AppBase()%>admin/songs.asp" style="display:inline">
                            <%=CsrfField()%>
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=rs("ID_SONG")%>">
                            <button class="icon-btn icon-btn--danger" type="submit" title="Eliminar"
                                    data-confirm="Se eliminara la canción y sus relaciones con artistas, álbumes y playlists."
                                    data-confirm-title="Eliminar canción" data-confirm-ok="Si, eliminar"><%=Icon("trash", "")%></button>
                        </form>
                    </div>
                </td>
            </tr>
            <% rs.MoveNext : Loop %>
        </tbody>
    </table>
</div>

<p class="small muted mt-16">Los archivos de audio se guardan físicamente en la carpeta <b>songFiles</b>; la base de datos solo almacena la ruta.</p>

<% rs.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
