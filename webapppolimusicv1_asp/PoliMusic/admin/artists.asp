<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Artistas"
pageSubtitle = "Administración del catálogo de artistas"
navActive = "admin_artists"
RequireAdmin()

Dim q, rs, action, idArtist, targetName, conn, errMsg, sql, fullName

action = FM("action")
errMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not CsrfValid() Then
        Call Flash("La sesión expiró. Vuelve a intentarlo.", "error")
        GoToPage AppBase() & "admin/artists.asp"
    ElseIf action = "delete" Then
        idArtist = ToInt(FM("id"), 0)
        targetName = DbScalar("SELECT ARTIST_NAME FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(idArtist)) & ""
        If idArtist <= 0 Or targetName = "" Then
            Call Flash("El artista indicado no existe.", "error")
        Else
            Set conn = Db()
            On Error Resume Next
            conn.BeginTrans
            Call DbExecute("DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(idArtist))
            Call DbExecute("DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(idArtist))
            If Err.Number <> 0 Then
                conn.RollbackTrans
                errMsg = "No se pudo eliminar el artista: " & Err.Description
                Err.Clear
            Else
                conn.CommitTrans
                Call Flash("Artista '" & targetName & "' eliminado.", "success")
            End If
            On Error GoTo 0
        End If
        GoToPage AppBase() & "admin/artists.asp"
    End If
End If

q = QS("q")
sql = "SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
      "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS SONGS, " & _
      "(SELECT COUNT(DISTINCT t.ID_ALBUM) FROM TBL_PLAYER p INNER JOIN TBL_TRACKLIST t ON t.ID_SONG = p.ID_SONG WHERE p.ID_ARTIST = ar.ID_ARTIST) AS ALBUMS " & _
      "FROM TBL_ARTIST ar "
If q <> "" Then
    sql = sql & "WHERE ar.ARTIST_NAME LIKE ? OR ar.ARTIST_LASTNAME LIKE ? ORDER BY ar.ARTIST_NAME"
    Set rs = DbQuery(sql, Array(LikeParam(q), LikeParam(q)))
Else
    sql = sql & "ORDER BY ar.ARTIST_NAME"
    Set rs = DbQuery(sql, Empty)
End If
%>
<!--#include file="../includes/layout_top.asp"-->

<% If errMsg <> "" Then %>
<div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
<% End If %>

<div class="admin-head">
    <div>
        <h2><%=rs.RecordCount%>&#32;<%=IIf(rs.RecordCount = 1, "artista", "artistas")%></h2>
        <p>La relacion canción-artista se almacena en TBL_PLAYER.</p>
    </div>
    <div class="row">
        <form method="get" action="<%=AppBase()%>admin/artists.asp" class="row" style="gap:6px">
            <input class="input" type="search" name="q" value="<%=HtmlAttr(q)%>" placeholder="Buscar artista" style="height:40px;min-width:190px">
            <button class="btn btn--ghost btn--sm" type="submit"><%=Icon("search", "")%></button>
        </form>
        <a class="btn btn--primary" href="<%=AppBase()%>admin/artist_form.asp"><%=Icon("plus", "")%> Nuevo artista</a>
    </div>
</div>

<div class="table-wrap">
    <table class="data">
        <thead>
            <tr>
                <th>ID</th>
                <th>Artista</th>
                <th>Apellido</th>
                <th class="num">Canciones</th>
                <th class="num">Álbumes</th>
                <th>Imagen</th>
                <th style="text-align:right">Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% Do Until rs.EOF %>
            <% fullName = Trim(rs("ARTIST_NAME") & " " & rs("ARTIST_LASTNAME") & "") %>
            <tr>
                <td class="num muted"><%=rs("ID_ARTIST")%></td>
                <td>
                    <div class="cell-user">
                        <%=Avatar(rs("ARTIST_IMAGE") & "", fullName, "avatar--sm")%>
                        <strong><%=HtmlEncode(rs("ARTIST_NAME") & "")%></strong>
                    </div>
                </td>
                <td class="small"><%=HtmlEncode(TextOr(rs("ARTIST_LASTNAME") & "", "-"))%></td>
                <td class="num"><%=ToInt(rs("SONGS"), 0)%></td>
                <td class="num"><%=ToInt(rs("ALBUMS"), 0)%></td>
                <td>
                    <% If AssetExists(rs("ARTIST_IMAGE") & "") Then %>
                    <span class="chip chip--admin">OK</span>
                    <% Else %>
                    <span class="chip chip--muted">Sin imagen</span>
                    <% End If %>
                </td>
                <td>
                    <div class="cell-actions">
                        <a class="icon-btn" href="<%=AppBase()%>admin/artist_form.asp?id=<%=rs("ID_ARTIST")%>" title="Editar"><%=Icon("edit", "")%></a>
                        <form method="post" action="<%=AppBase()%>admin/artists.asp" style="display:inline">
                            <%=CsrfField()%>
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=rs("ID_ARTIST")%>">
                            <button class="icon-btn icon-btn--danger" type="submit" title="Eliminar"
                                    data-confirm="Se eliminara el artista y sus relaciones con las canciones."
                                    data-confirm-title="Eliminar artista" data-confirm-ok="Si, eliminar"><%=Icon("trash", "")%></button>
                        </form>
                    </div>
                </td>
            </tr>
            <% rs.MoveNext : Loop %>
        </tbody>
    </table>
</div>

<% rs.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
