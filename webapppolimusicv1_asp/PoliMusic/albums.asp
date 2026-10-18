<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Álbumes"
pageSubtitle = "Colecciones que agrupan las canciones del catálogo"
navActive = "albums"
RequireLogin()

Dim q, sortBy, rs, total, orderBy
q = QS("q")
sortBy = LCase(QS("sort"))
If sortBy <> "year" And sortBy <> "title" And sortBy <> "tracks" Then sortBy = "year"

Select Case sortBy
    Case "title"  : orderBy = "a.TITLE_ALBUM ASC"
    Case "tracks" : orderBy = "TRACKS DESC, a.TITLE_ALBUM ASC"
    Case Else     : orderBy = "a.RELEASE_YEAR DESC, a.TITLE_ALBUM ASC"
End Select

Dim sql
sql = "SELECT a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
      "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = a.ID_ALBUM) AS TRACKS FROM TBL_ALBUM a "

If q <> "" Then
    sql = sql & "WHERE a.TITLE_ALBUM LIKE ? ORDER BY " & orderBy
    Set rs = DbQuery(sql, Array(LikeParam(q)))
Else
    sql = sql & "ORDER BY " & orderBy
    Set rs = DbQuery(sql, Empty)
End If

total = 0
%>
<!--#include file="includes/layout_top.asp"-->

<form class="filters" method="get" action="<%=AppBase()%>albums.asp">
    <input class="input filters__grow" type="search" name="q" value="<%=HtmlAttr(q)%>" placeholder="Buscar álbum por título" aria-label="Buscar álbumes">
    <select class="select" name="sort" aria-label="Ordenar álbumes">
        <option value="year"<%=IIf(sortBy = "year", " selected", "")%>>Mas recientes</option>
        <option value="title"<%=IIf(sortBy = "title", " selected", "")%>>Por título</option>
        <option value="tracks"<%=IIf(sortBy = "tracks", " selected", "")%>>Mas canciones</option>
    </select>
    <button class="btn btn--ghost btn--sm" type="submit"><%=Icon("search", "")%> Buscar</button>
    <% If q <> "" Then %><a class="btn btn--ghost btn--sm" href="<%=AppBase()%>albums.asp">Limpiar</a><% End If %>
</form>

<% If rs.EOF Then %>
    <%=EmptyState("disc", "Sin álbumes", "No hay álbumes que coincidan con la busqueda.", "<a class=""btn btn--ghost btn--sm"" href=""" & AppBase() & "albums.asp"">Ver todos</a>")%>
<% Else %>
<div class="grid grid--albums">
    <% Do Until rs.EOF %>
    <%=AlbumTile(rs("ID_ALBUM"), rs("TITLE_ALBUM") & "", rs("RELEASE_YEAR"), rs("ALBUM_IMAGE_PATH") & "", rs("TRACKS"))%>
    <% total = total + 1 %>
    <% rs.MoveNext : Loop %>
</div>
<p class="small muted mt-16"><b class="mono"><%=total%></b> <%=IIf(total = 1, "álbum", "álbumes")%> en el catálogo.</p>
<% End If %>

<% rs.Close %>
<!--#include file="includes/layout_bottom.asp"-->
