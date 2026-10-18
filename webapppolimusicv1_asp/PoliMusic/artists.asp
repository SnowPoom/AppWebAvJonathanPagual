<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Artistas"
pageSubtitle = "Las voces que dan vida al catálogo"
navActive = "artists"
RequireLogin()

Dim q, rs, total
q = QS("q")

If q <> "" Then
    Set rs = DbQuery("SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS SONGS " & _
        "FROM TBL_ARTIST ar WHERE ar.ARTIST_NAME LIKE ? OR ar.ARTIST_LASTNAME LIKE ? " & _
        "ORDER BY ar.ARTIST_NAME", Array(LikeParam(q), LikeParam(q)))
Else
    Set rs = DbQuery("SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS SONGS " & _
        "FROM TBL_ARTIST ar ORDER BY ar.ARTIST_NAME", Empty)
End If

total = 0
%>
<!--#include file="includes/layout_top.asp"-->

<form class="filters" method="get" action="<%=AppBase()%>artists.asp">
    <input class="input filters__grow" type="search" name="q" value="<%=HtmlAttr(q)%>" placeholder="Buscar artista por nombre o apellido" aria-label="Buscar artistas">
    <button class="btn btn--ghost btn--sm" type="submit"><%=Icon("search", "")%> Buscar</button>
    <% If q <> "" Then %><a class="btn btn--ghost btn--sm" href="<%=AppBase()%>artists.asp">Limpiar</a><% End If %>
</form>

<% If rs.EOF Then %>
    <%=EmptyState("mic", "Sin artistas", "No encontramos artistas que coincidan con la busqueda.", "<a class=""btn btn--ghost btn--sm"" href=""" & AppBase() & "artists.asp"">Ver todos</a>")%>
<% Else %>
<div class="grid grid--artists">
    <% Do Until rs.EOF %>
    <%=ArtistTile(rs("ID_ARTIST"), Trim(rs("ARTIST_NAME") & " " & rs("ARTIST_LASTNAME") & ""), rs("ARTIST_IMAGE") & "", rs("SONGS"))%>
    <% total = total + 1 %>
    <% rs.MoveNext : Loop %>
</div>
<p class="small muted mt-16"><b class="mono"><%=total%></b> <%=IIf(total = 1, "artista", "artistas")%> en el catálogo.</p>
<% End If %>

<% rs.Close %>
<!--#include file="includes/layout_bottom.asp"-->
