<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Buscar"
pageSubtitle = "Resultados en todo el catálogo"
navActive = "search"
RequireLogin()

Dim q, likeValue, rsSongs, rsArtists, rsAlbums, idx, totalHits
q = QS("q")
likeValue = LikeParam(q)
totalHits = 0

If q <> "" Then
    Set rsSongs = DbQuery("SELECT TOP 20 * FROM (" & SongBaseSql() & _
        " WHERE s.SONG_NAME LIKE ? OR EXISTS (SELECT 1 FROM TBL_PLAYER pp INNER JOIN TBL_ARTIST aa ON aa.ID_ARTIST = pp.ID_ARTIST " & _
        "WHERE pp.ID_SONG = s.ID_SONG AND (aa.ARTIST_NAME LIKE ? OR aa.ARTIST_LASTNAME LIKE ?)) " & _
        "OR EXISTS (SELECT 1 FROM TBL_TRACKLIST tt INNER JOIN TBL_ALBUM ab ON ab.ID_ALBUM = tt.ID_ALBUM " & _
        "WHERE tt.ID_SONG = s.ID_SONG AND ab.TITLE_ALBUM LIKE ?)) q ORDER BY q.PLAYS DESC, q.SONG_NAME", _
        Array(likeValue, likeValue, likeValue, likeValue))

    Set rsArtists = DbQuery("SELECT ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS SONGS " & _
        "FROM TBL_ARTIST ar WHERE ar.ARTIST_NAME LIKE ? OR ar.ARTIST_LASTNAME LIKE ? ORDER BY ar.ARTIST_NAME", _
        Array(likeValue, likeValue))

    Set rsAlbums = DbQuery("SELECT TOP 12 a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
        "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = a.ID_ALBUM) AS TRACKS " & _
        "FROM TBL_ALBUM a WHERE a.TITLE_ALBUM LIKE ? ORDER BY a.RELEASE_YEAR DESC", Array(likeValue))

    totalHits = rsSongs.RecordCount + rsArtists.RecordCount + rsAlbums.RecordCount
End If
%>
<!--#include file="includes/layout_top.asp"-->

<form class="filters" method="get" action="<%=AppBase()%>search.asp">
    <input class="input filters__grow" type="search" name="q" value="<%=HtmlAttr(q)%>" placeholder="Que quieres escuchar hoy?" aria-label="Buscar en PoliMusic" autofocus>
    <button class="btn btn--primary btn--sm" type="submit"><%=Icon("search", "")%> Buscar</button>
</form>

<% If q = "" Then %>
    <%=EmptyState("compass", "Empieza a buscar", "Escribe el nombre de una canción, un artista o un álbum para ver los resultados.", "<a class=""btn btn--ghost btn--sm"" href=""" & AppBase() & "songs.asp"">Explorar catálogo</a>")%>
<% Else %>
    <p class="small muted" style="margin-bottom:18px">
        <b class="mono"><%=totalHits%></b> resultados para <b style="color:#e8f4ff">&ldquo;<%=HtmlEncode(q)%>&rdquo;</b>
    </p>

    <section class="section" style="margin-top:0">
        <div class="section__head">
            <h2 class="section__title"><%=Icon("music", "")%> Canciones <span class="chip chip--muted"><%=rsSongs.RecordCount%></span></h2>
            <a class="section__link" href="<%=AppBase()%>songs.asp?q=<%=Server.URLEncode(q)%>">Ver en el catálogo <%=Icon("chevron-right", "")%></a>
        </div>
        <% If rsSongs.EOF Then %>
            <p class="small muted">Sin coincidencias en canciones.</p>
        <% Else %>
        <div class="panel panel--flush">
            <div class="songlist">
                <%
                idx = 1
                Do Until rsSongs.EOF
                    Response.Write SongRowHtml(rsSongs("ID_SONG"), rsSongs("SONG_NAME") & "", rsSongs("SONG_PATH") & "", rsSongs("PLAYS"), _
                        rsSongs("ARTISTS") & "", rsSongs("ALBUM_TITLE") & "", rsSongs("ALBUM_IMAGE") & "", "", idx)
                    idx = idx + 1
                    rsSongs.MoveNext
                Loop
                %>
            </div>
        </div>
        <% End If %>
    </section>

    <section class="section">
        <div class="section__head">
            <h2 class="section__title"><%=Icon("mic", "")%> Artistas <span class="chip chip--muted"><%=rsArtists.RecordCount%></span></h2>
        </div>
        <% If rsArtists.EOF Then %>
            <p class="small muted">Sin coincidencias en artistas.</p>
        <% Else %>
        <div class="grid grid--artists">
            <% Do Until rsArtists.EOF %>
            <%=ArtistTile(rsArtists("ID_ARTIST"), Trim(rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME") & ""), rsArtists("ARTIST_IMAGE") & "", rsArtists("SONGS"))%>
            <% rsArtists.MoveNext : Loop %>
        </div>
        <% End If %>
    </section>

    <section class="section">
        <div class="section__head">
            <h2 class="section__title"><%=Icon("disc", "")%> Álbumes <span class="chip chip--muted"><%=rsAlbums.RecordCount%></span></h2>
        </div>
        <% If rsAlbums.EOF Then %>
            <p class="small muted">Sin coincidencias en álbumes.</p>
        <% Else %>
        <div class="grid grid--albums">
            <% Do Until rsAlbums.EOF %>
            <%=AlbumTile(rsAlbums("ID_ALBUM"), rsAlbums("TITLE_ALBUM") & "", rsAlbums("RELEASE_YEAR"), rsAlbums("ALBUM_IMAGE_PATH") & "", rsAlbums("TRACKS"))%>
            <% rsAlbums.MoveNext : Loop %>
        </div>
        <% End If %>
    </section>
<% End If %>

<% If q <> "" Then rsSongs.Close : rsArtists.Close : rsAlbums.Close %>
<!--#include file="includes/layout_bottom.asp"-->
