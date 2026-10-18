<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
RequireLogin()

Dim idArtist, rsArtist, rsSongs, rsAlbums, fullName, idx, songCount

idArtist = ToInt(QS("id"), 0)
If idArtist <= 0 Then
    Flash "Artista no especificado.", "error"
    GoToPage AppBase() & "artists.asp"
End If

Set rsArtist = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(idArtist))
If rsArtist.EOF Then
    rsArtist.Close
    Flash "El artista solicitado no existe.", "error"
    GoToPage AppBase() & "artists.asp"
End If

fullName = Trim(rsArtist("ARTIST_NAME") & " " & rsArtist("ARTIST_LASTNAME") & "")
pageTitle = fullName
pageSubtitle = "Perfil del artista"
navActive = "artists"

Set rsSongs = DbQuery("SELECT * FROM (" & SongBaseSql() & _
    " WHERE EXISTS (SELECT 1 FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG AND p.ID_ARTIST = ?)) q " & _
    "ORDER BY q.PLAYS DESC, q.SONG_NAME", Array(idArtist))

Set rsAlbums = DbQuery("SELECT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = al.ID_ALBUM) AS TRACKS " & _
    "FROM TBL_ALBUM al WHERE EXISTS (SELECT 1 FROM TBL_TRACKLIST t INNER JOIN TBL_PLAYER p ON p.ID_SONG = t.ID_SONG " & _
    "WHERE t.ID_ALBUM = al.ID_ALBUM AND p.ID_ARTIST = ?) ORDER BY al.RELEASE_YEAR DESC", Array(idArtist))

songCount = 0
%>
<!--#include file="includes/layout_top.asp"-->

<a class="btn btn--ghost btn--sm" href="<%=AppBase()%>artists.asp" style="margin-bottom:16px"><%=Icon("arrow-left", "")%> Volver a artistas</a>

<section class="detail">
    <div class="detail__media" style="width:190px;flex-basis:190px">
        <%=Cover(rsArtist("ARTIST_IMAGE") & "", fullName, "cover--rounded")%>
    </div>
    <div class="detail__body">
        <div class="detail__kind">Artista</div>
        <h2><%=HtmlEncode(fullName)%></h2>
        <div class="detail__meta">
            <span class="chip"><%=Icon("music", "")%>&#32;<%=rsSongs.RecordCount%>&#32;<%=IIf(rsSongs.RecordCount = 1, "canción", "canciones")%></span>
            <span class="dot"></span>
            <span><%=rsAlbums.RecordCount%>&#32;<%=IIf(rsAlbums.RecordCount = 1, "álbum relacionado", "álbumes relacionados")%></span>
        </div>
        <div class="detail__actions">
            <button class="btn btn--primary" type="button" data-play-all>
                <%=Icon("play", "")%> Reproducir todo
            </button>
        </div>
    </div>
</section>

<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("music", "")%> Canciones</h2>
    </div>
    <% If rsSongs.EOF Then %>
        <%=EmptyState("music", "Sin canciones", "Este artista aún no tiene canciones asociadas en el catálogo.", "")%>
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

<% If Not rsAlbums.EOF Then %>
<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("disc", "")%> Álbumes relacionados</h2>
    </div>
    <div class="grid grid--albums">
        <% Do Until rsAlbums.EOF %>
        <%=AlbumTile(rsAlbums("ID_ALBUM"), rsAlbums("TITLE_ALBUM") & "", rsAlbums("RELEASE_YEAR"), rsAlbums("ALBUM_IMAGE_PATH") & "", rsAlbums("TRACKS"))%>
        <% rsAlbums.MoveNext : Loop %>
    </div>
</section>
<% End If %>

<% rsArtist.Close : rsSongs.Close : rsAlbums.Close %>
<!--#include file="includes/layout_bottom.asp"-->
