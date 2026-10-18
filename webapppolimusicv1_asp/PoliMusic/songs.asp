<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Canciones"
pageSubtitle = "Explora el catálogo completo de PoliMusic"
navActive = "songs"
RequireLogin()

Dim q, fArtist, fAlbum, sortBy, page, pageSize, totalRows, totalPages
Dim whereSql, conds(8), cCount, prm(), pCount, i, prmArr, dataPrm()
Dim rs, rsArtists, rsAlbums, orderBy, baseQs, idx, pStart, pEnd, offsetRows

q = QS("q")
fArtist = ToInt(QS("artist"), 0)
fAlbum = ToInt(QS("álbum"), 0)
sortBy = LCase(QS("sort"))
If sortBy <> "name" And sortBy <> "artist" And sortBy <> "plays" Then sortBy = "plays"
page = ToInt(QS("page"), 1)
If page < 1 Then page = 1
pageSize = 10

ReDim prm(24)
pCount = 0
cCount = 0

If q <> "" Then
    conds(cCount) = "(s.SONG_NAME LIKE ? OR EXISTS (SELECT 1 FROM TBL_PLAYER pp INNER JOIN TBL_ARTIST aa ON aa.ID_ARTIST = pp.ID_ARTIST " & _
                   "WHERE pp.ID_SONG = s.ID_SONG AND (aa.ARTIST_NAME LIKE ? OR aa.ARTIST_LASTNAME LIKE ?)) " & _
                   "OR EXISTS (SELECT 1 FROM TBL_TRACKLIST tt INNER JOIN TBL_ALBUM ab ON ab.ID_ALBUM = tt.ID_ALBUM " & _
                   "WHERE tt.ID_SONG = s.ID_SONG AND ab.TITLE_ALBUM LIKE ?))"
    cCount = cCount + 1
    prm(pCount) = LikeParam(q) : pCount = pCount + 1
    prm(pCount) = LikeParam(q) : pCount = pCount + 1
    prm(pCount) = LikeParam(q) : pCount = pCount + 1
    prm(pCount) = LikeParam(q) : pCount = pCount + 1
End If

If fArtist > 0 Then
    conds(cCount) = "EXISTS (SELECT 1 FROM TBL_PLAYER p2 WHERE p2.ID_SONG = s.ID_SONG AND p2.ID_ARTIST = ?)"
    cCount = cCount + 1
    prm(pCount) = fArtist : pCount = pCount + 1
End If

If fAlbum > 0 Then
    conds(cCount) = "EXISTS (SELECT 1 FROM TBL_TRACKLIST t2 WHERE t2.ID_SONG = s.ID_SONG AND t2.ID_ALBUM = ?)"
    cCount = cCount + 1
    prm(pCount) = fAlbum : pCount = pCount + 1
End If

whereSql = ""
For i = 0 To cCount - 1
    If whereSql = "" Then
        whereSql = " WHERE " & conds(i)
    Else
        whereSql = whereSql & " AND " & conds(i)
    End If
Next

If pCount > 0 Then
    ReDim Preserve prm(pCount - 1)
    prmArr = prm
Else
    prmArr = Empty
End If

totalRows = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG s" & whereSql, prmArr), 0)
totalPages = Int((totalRows + pageSize - 1) / pageSize)
If totalPages < 1 Then totalPages = 1
If page > totalPages Then page = totalPages

Select Case sortBy
    Case "name"   : orderBy = "q.SONG_NAME ASC, q.ID_SONG ASC"
    Case "artist" : orderBy = "q.ARTISTS ASC, q.SONG_NAME ASC"
    Case Else     : orderBy = "q.PLAYS DESC, q.SONG_NAME ASC"
End Select

offsetRows = (page - 1) * pageSize
ReDim dataPrm(pCount + 1)
For i = 0 To pCount - 1
    dataPrm(i) = prm(i)
Next
dataPrm(pCount) = offsetRows
dataPrm(pCount + 1) = pageSize

Set rs = DbQuery("SELECT * FROM (" & SongBaseSql() & whereSql & ") q ORDER BY " & orderBy & _
                 " OFFSET ? ROWS FETCH NEXT ? ROWS ONLY", dataPrm)

Set rsArtists = DbQuery("SELECT ar.ID_ARTIST, LTRIM(RTRIM(ar.ARTIST_NAME + ' ' + ISNULL(ar.ARTIST_LASTNAME,''))) AS FULLNAME " & _
                        "FROM TBL_ARTIST ar ORDER BY FULLNAME", Empty)
Set rsAlbums = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM FROM TBL_ALBUM ORDER BY TITLE_ALBUM", Empty)

baseQs = "&q=" & Server.URLEncode(q) & "&artist=" & fArtist & "&álbum=" & fAlbum & "&sort=" & sortBy
%>
<!--#include file="includes/layout_top.asp"-->

<form class="filters" method="get" action="<%=AppBase()%>songs.asp">
    <input class="input filters__grow" type="search" name="q" value="<%=HtmlAttr(q)%>"
           placeholder="Buscar por canción, artista o álbum" aria-label="Buscar canciones">
    <select class="select" name="artist" aria-label="Filtrar por artista">
        <option value="0">Todos los artistas</option>
        <% Do Until rsArtists.EOF %>
        <option value="<%=rsArtists("ID_ARTIST")%>"<%=IIf(fArtist = ToInt(rsArtists("ID_ARTIST"), 0), " selected", "")%>><%=HtmlEncode(rsArtists("FULLNAME"))%></option>
        <% rsArtists.MoveNext : Loop %>
    </select>
    <select class="select" name="álbum" aria-label="Filtrar por álbum">
        <option value="0">Todos los álbumes</option>
        <% Do Until rsAlbums.EOF %>
        <option value="<%=rsAlbums("ID_ALBUM")%>"<%=IIf(fAlbum = ToInt(rsAlbums("ID_ALBUM"), 0), " selected", "")%>><%=HtmlEncode(rsAlbums("TITLE_ALBUM"))%></option>
        <% rsAlbums.MoveNext : Loop %>
    </select>
    <input type="hidden" name="sort" value="<%=HtmlAttr(sortBy)%>">
    <button class="btn btn--ghost btn--sm" type="submit"><%=Icon("search", "")%> Filtrar</button>
    <% If q <> "" Or fArtist > 0 Or fAlbum > 0 Then %>
    <a class="btn btn--ghost btn--sm" href="<%=AppBase()%>songs.asp">Limpiar</a>
    <% End If %>
</form>

<div class="row row--between row--wrap" style="margin-bottom:14px">
    <p class="small muted"><b class="mono"><%=totalRows%></b> <%=IIf(totalRows = 1, "canción encontrada", "canciones encontradas")%></p>
    <div class="seg">
        <a class="<%=IIf(sortBy = "plays", "is-active", "")%>" href="<%=AppBase()%>songs.asp?page=1&q=<%=Server.URLEncode(q)%>&artist=<%=fArtist%>&álbum=<%=fAlbum%>&sort=plays">Más escuchadas</a>
        <a class="<%=IIf(sortBy = "name", "is-active", "")%>" href="<%=AppBase()%>songs.asp?page=1&q=<%=Server.URLEncode(q)%>&artist=<%=fArtist%>&álbum=<%=fAlbum%>&sort=name">Por título</a>
        <a class="<%=IIf(sortBy = "artist", "is-active", "")%>" href="<%=AppBase()%>songs.asp?page=1&q=<%=Server.URLEncode(q)%>&artist=<%=fArtist%>&álbum=<%=fAlbum%>&sort=artist">Por artista</a>
    </div>
</div>

<% If rs.EOF Then %>
    <%=EmptyState("music", "Sin resultados", "No encontramos canciones con esos criterios. Prueba con otra busqueda o limpia los filtros.", "<a class=""btn btn--ghost btn--sm"" href=""" & AppBase() & "songs.asp"">Ver todo el catálogo</a>")%>
<% Else %>
<div class="panel panel--flush">
    <div class="songlist">
        <%
        idx = offsetRows + 1
        Do Until rs.EOF
            Response.Write SongRowHtml(rs("ID_SONG"), rs("SONG_NAME") & "", rs("SONG_PATH") & "", rs("PLAYS"), _
                rs("ARTISTS") & "", rs("ALBUM_TITLE") & "", rs("ALBUM_IMAGE") & "", "", idx)
            idx = idx + 1
            rs.MoveNext
        Loop
        %>
    </div>
</div>

<% If totalPages > 1 Then %>
<nav class="pager" aria-label="Paginacion">
    <% If page > 1 Then %>
    <a href="<%=AppBase()%>songs.asp?page=<%=page-1%><%=baseQs%>" aria-label="Página anterior">&laquo;</a>
    <% End If %>
    <%
    pStart = page - 2
    If pStart < 1 Then pStart = 1
    pEnd = pStart + 4
    If pEnd > totalPages Then
        pEnd = totalPages
        pStart = pEnd - 4
        If pStart < 1 Then pStart = 1
    End If
    For i = pStart To pEnd
    %>
    <% If i = page Then %>
    <span class="is-current"><%=i%></span>
    <% Else %>
    <a href="<%=AppBase()%>songs.asp?page=<%=i%><%=baseQs%>"><%=i%></a>
    <% End If %>
    <% Next %>
    <% If page < totalPages Then %>
    <a href="<%=AppBase()%>songs.asp?page=<%=page+1%><%=baseQs%>" aria-label="Página siguiente">&raquo;</a>
    <% End If %>
    <span class="muted" style="border:0;background:none"><%=page%> / <%=totalPages%></span>
</nav>
<% End If %>
<% End If %>

<% rs.Close : rsArtists.Close : rsAlbums.Close %>
<!--#include file="includes/layout_bottom.asp"-->
