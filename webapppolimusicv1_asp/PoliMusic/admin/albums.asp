<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Álbumes"
pageSubtitle = "Administración de álbumes y sus listas de canciones"
navActive = "admin_albums"
RequireAdmin()

Dim q, rs, action, idAlbum, targetName, conn, errMsg, sql

action = FM("action")
errMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not CsrfValid() Then
        Call Flash("La sesión expiró. Vuelve a intentarlo.", "error")
        GoToPage AppBase() & "admin/albums.asp"
    ElseIf action = "delete" Then
        idAlbum = ToInt(FM("id"), 0)
        targetName = DbScalar("SELECT TITLE_ALBUM FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(idAlbum)) & ""
        If idAlbum <= 0 Or targetName = "" Then
            Call Flash("El álbum indicado no existe.", "error")
        Else
            Set conn = Db()
            On Error Resume Next
            conn.BeginTrans
            Call DbExecute("DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(idAlbum))
            Call DbExecute("DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(idAlbum))
            If Err.Number <> 0 Then
                conn.RollbackTrans
                errMsg = "No se pudo eliminar el álbum: " & Err.Description
                Err.Clear
            Else
                conn.CommitTrans
                Call Flash("Álbum '" & targetName & "' eliminado.", "success")
            End If
            On Error GoTo 0
        End If
        GoToPage AppBase() & "admin/albums.asp"
    End If
End If

q = QS("q")
sql = "SELECT a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
      "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = a.ID_ALBUM) AS TRACKS, " & _
      "(SELECT COUNT(*) FROM TBL_TRACKLIST t INNER JOIN TBL_PLAYER p ON p.ID_SONG = t.ID_SONG WHERE t.ID_ALBUM = a.ID_ALBUM) AS LINKS " & _
      "FROM TBL_ALBUM a "
If q <> "" Then
    sql = sql & "WHERE a.TITLE_ALBUM LIKE ? ORDER BY a.TITLE_ALBUM"
    Set rs = DbQuery(sql, Array(LikeParam(q)))
Else
    sql = sql & "ORDER BY a.RELEASE_YEAR DESC, a.TITLE_ALBUM"
    Set rs = DbQuery(sql, Empty)
End If
%>
<!--#include file="../includes/layout_top.asp"-->

<% If errMsg <> "" Then %>
<div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
<% End If %>

<div class="admin-head">
    <div>
        <h2><%=rs.RecordCount%>&#32;<%=IIf(rs.RecordCount = 1, "álbum", "álbumes")%></h2>
        <p>Las pistas se relacionan mediante TBL_TRACKLIST.</p>
    </div>
    <div class="row">
        <form method="get" action="<%=AppBase()%>admin/albums.asp" class="row" style="gap:6px">
            <input class="input" type="search" name="q" value="<%=HtmlAttr(q)%>" placeholder="Buscar álbum" style="height:40px;min-width:190px">
            <button class="btn btn--ghost btn--sm" type="submit"><%=Icon("search", "")%></button>
        </form>
        <a class="btn btn--primary" href="<%=AppBase()%>admin/album_form.asp"><%=Icon("plus", "")%> Nuevo álbum</a>
    </div>
</div>

<div class="table-wrap">
    <table class="data">
        <thead>
            <tr>
                <th>ID</th>
                <th>Álbum</th>
                <th class="num">Año</th>
                <th class="num">Canciones</th>
                <th>Portada</th>
                <th style="text-align:right">Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% Do Until rs.EOF %>
            <tr>
                <td class="num muted"><%=rs("ID_ALBUM")%></td>
                <td>
                    <div class="cell-user">
                        <%=Cover(rs("ALBUM_IMAGE_PATH") & "", rs("TITLE_ALBUM") & "", "cover--thumb")%>
                        <div>
                            <strong><%=HtmlEncode(rs("TITLE_ALBUM") & "")%></strong>
                            <div class="small muted mono"><%=HtmlEncode(rs("ALBUM_IMAGE_PATH") & "")%></div>
                        </div>
                    </div>
                </td>
                <td class="num"><%=IIf(ToInt(rs("RELEASE_YEAR"), 0) > 0, rs("RELEASE_YEAR"), "-")%></td>
                <td class="num"><%=ToInt(rs("TRACKS"), 0)%></td>
                <td>
                    <% If AssetExists(rs("ALBUM_IMAGE_PATH") & "") Then %>
                    <span class="chip chip--admin">OK</span>
                    <% Else %>
                    <span class="chip chip--muted">Sin portada</span>
                    <% End If %>
                </td>
                <td>
                    <div class="cell-actions">
                        <a class="icon-btn" href="<%=AppBase()%>album.asp?id=<%=rs("ID_ALBUM")%>" title="Ver en el sitio"><%=Icon("eye", "")%></a>
                        <a class="icon-btn" href="<%=AppBase()%>admin/album_form.asp?id=<%=rs("ID_ALBUM")%>" title="Editar"><%=Icon("edit", "")%></a>
                        <form method="post" action="<%=AppBase()%>admin/albums.asp" style="display:inline">
                            <%=CsrfField()%>
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=rs("ID_ALBUM")%>">
                            <button class="icon-btn icon-btn--danger" type="submit" title="Eliminar"
                                    data-confirm="Se eliminara el álbum y su lista de canciones (las canciones no se borran)."
                                    data-confirm-title="Eliminar álbum" data-confirm-ok="Si, eliminar"><%=Icon("trash", "")%></button>
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
