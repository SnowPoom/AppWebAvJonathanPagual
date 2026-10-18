<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Inicio"
pageSubtitle = "Tu resumen musical de hoy"
navActive = "home"
RequireLogin()

Dim saludo, totalSongs, totalArtists, totalAlbums, totalMine, totalPlays
Dim rsSongs, rsAlbums, rsArtists, rsPl, idx, i, firstName

If Hour(Now()) < 12 Then
    saludo = "Buenos días"
ElseIf Hour(Now()) < 19 Then
    saludo = "Buenas tardes"
Else
    saludo = "Buenas noches"
End If

firstName = Session("Usuario") & ""
If InStr(firstName, ".") > 0 Then firstName = Left(firstName, InStr(firstName, ".") - 1)
If InStr(firstName, " ") > 0 Then firstName = Left(firstName, InStr(firstName, " ") - 1)

totalSongs = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG", Empty), 0)
totalArtists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", Empty), 0)
totalAlbums = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM", Empty), 0)
totalMine = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(CurrentUserId())), 0)
totalPlays = ToInt(DbScalar("SELECT SUM(ISNULL(PLAYS,0)) FROM TBL_SONG", Empty), 0)

Set rsSongs = DbQuery("SELECT TOP 6 q.* FROM (" & SongBaseSql() & ") q ORDER BY q.PLAYS DESC, q.ID_SONG ASC", Empty)
Set rsAlbums = DbQuery("SELECT TOP 5 a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = a.ID_ALBUM) AS TRACKS " & _
    "FROM TBL_ALBUM a ORDER BY a.RELEASE_YEAR DESC, a.ID_ALBUM", Empty)
Set rsArtists = DbQuery("SELECT TOP 5 ar.ID_ARTIST, ar.ARTIST_NAME, ar.ARTIST_LASTNAME, ar.ARTIST_IMAGE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = ar.ID_ARTIST) AS SONGS " & _
    "FROM TBL_ARTIST ar ORDER BY SONGS DESC, ar.ARTIST_NAME", Empty)
Set rsPl = DbQuery("SELECT TOP 4 p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, p.CREATION_DATE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONGS " & _
    "FROM TBL_PLAYLIST p WHERE p.ID_USER = ? ORDER BY p.CREATION_DATE DESC, p.ID_PLAYLIST DESC", Array(CurrentUserId()))
%>
<!--#include file="includes/layout_top.asp"-->

<section class="hero">
    <div class="hero__art" aria-hidden="true">
        <svg viewBox="0 0 400 400" fill="none">
            <defs>
                <linearGradient id="hgrad" x1="0" y1="0" x2="1" y2="1">
                    <stop offset="0" stop-color="#5eead4"/><stop offset=".5" stop-color="#22d3ee"/><stop offset="1" stop-color="#3b82f6"/>
                </linearGradient>
                <radialGradient id="hcore" cx=".35" cy=".3" r=".8">
                    <stop offset="0" stop-color="#6ee7b7"/><stop offset=".6" stop-color="#22d3ee"/><stop offset="1" stop-color="#2563eb"/>
                </radialGradient>
            </defs>
            <circle cx="200" cy="200" r="186" stroke="url(#hgrad)" stroke-width="1" opacity=".35"/>
            <circle cx="200" cy="200" r="162" stroke="url(#hgrad)" stroke-width="1" opacity=".28"/>
            <circle cx="200" cy="200" r="138" stroke="url(#hgrad)" stroke-width="1" opacity=".22"/>
            <circle cx="200" cy="200" r="114" stroke="url(#hgrad)" stroke-width="1" opacity=".18"/>
            <circle cx="200" cy="200" r="88" fill="url(#hcore)" opacity=".55"/>
            <circle cx="200" cy="200" r="60" stroke="#04121f" stroke-width="1.5" opacity=".5"/>
            <circle cx="200" cy="200" r="16" fill="#04121f"/>
            <circle cx="200" cy="200" r="16" stroke="url(#hgrad)" stroke-width="1.5"/>
            <path d="M40 200h34l14-30 20 62 18-96 20 118 18-72 16 38 14-20h150" stroke="url(#hgrad)" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" opacity=".75"/>
        </svg>
    </div>
    <div class="hero__inner">
        <span class="hero__eyebrow"><%=Icon("sparkles", "")%><%=saludo%> &middot; sesión <%=RoleLabel(Session("UsuarioTipo"))%></span>
        <h2>Bienvenido, <span class="grad"><%=HtmlEncode(Session("Usuario") & "")%></span>.<br>Tu música te está esperando.</h2>
        <p>Tu biblioteca tiene <%=totalSongs%> canciones, <%=totalArtists%> artistas y <%=totalAlbums%> álbumes listos para sonar. Arma una playlist o continua donde lo dejaste.</p>
        <div class="hero__actions">
            <a class="btn btn--primary" href="<%=AppBase()%>songs.asp"><%=Icon("play", "")%> Explorar canciones</a>
            <a class="btn btn--ghost" href="<%=AppBase()%>playlists.asp"><%=Icon("list", "")%> Mis playlists</a>
        </div>
    </div>
</section>

<div class="grid grid--stats mt-24">
    <div class="stat">
        <span class="stat__ico"><%=Icon("music", "")%></span>
        <div><b><%=totalSongs%></b><span>Canciones</span></div>
    </div>
    <div class="stat">
        <span class="stat__ico"><%=Icon("mic", "")%></span>
        <div><b><%=totalArtists%></b><span>Artistas</span></div>
    </div>
    <div class="stat">
        <span class="stat__ico"><%=Icon("disc", "")%></span>
        <div><b><%=totalAlbums%></b><span>Álbumes</span></div>
    </div>
    <div class="stat">
        <span class="stat__ico"><%=Icon("list", "")%></span>
        <div><b><%=totalMine%></b><span>Mis playlists</span></div>
    </div>
</div>

<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("list", "")%> Tus playlists</h2>
        <a class="section__link" href="<%=AppBase()%>playlists.asp">Gestionar <%=Icon("chevron-right", "")%></a>
    </div>
    <% If rsPl.EOF Then %>
        <%=EmptyState("list", "Todavía no tienes playlists", "Crea una playlist para agrupar las canciones que mas te gustan.", "<a class=""btn btn--primary btn--sm"" href=""" & AppBase() & "playlists.asp?nueva=1"">Crear playlist</a>")%>
    <% Else %>
    <div class="grid grid--playlists">
        <% Do Until rsPl.EOF %>
        <%=PlaylistTile(rsPl("ID_PLAYLIST"), rsPl("TITLE") & "", rsPl("TYPE"), rsPl("PHOTO") & "", rsPl("SONGS"), rsPl("CREATION_DATE"), Session("Usuario") & "")%>
        <% rsPl.MoveNext : Loop %>
    </div>
    <% End If %>
</section>

<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("wave", "")%> Populares en PoliMusic</h2>
        <a class="section__link" href="<%=AppBase()%>songs.asp">Ver todas <%=Icon("chevron-right", "")%></a>
    </div>
    <div class="panel panel--flush">
        <div class="songlist">
            <%
            idx = 1
            Do Until rsSongs.EOF
                Response.Write SongRowHtml(rsSongs("ID_SONG"), rsSongs("SONG_NAME") & "", rsSongs("SONG_PATH") & "", _
                    rsSongs("PLAYS"), rsSongs("ARTISTS") & "", rsSongs("ALBUM_TITLE") & "", rsSongs("ALBUM_IMAGE") & "", "", idx)
                idx = idx + 1
                rsSongs.MoveNext
            Loop
            %>
        </div>
    </div>
    <p class="small muted mt-8">Reproducciones acumuladas en el catálogo: <b class="mono"><%=totalPlays%></b></p>
</section>

<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("disc", "")%> Álbumes destacados</h2>
        <a class="section__link" href="<%=AppBase()%>albums.asp">Ver todos <%=Icon("chevron-right", "")%></a>
    </div>
    <div class="grid grid--albums">
        <% Do Until rsAlbums.EOF %>
        <%=AlbumTile(rsAlbums("ID_ALBUM"), rsAlbums("TITLE_ALBUM") & "", rsAlbums("RELEASE_YEAR"), rsAlbums("ALBUM_IMAGE_PATH") & "", rsAlbums("TRACKS"))%>
        <% rsAlbums.MoveNext : Loop %>
    </div>
</section>

<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("mic", "")%> Artistas del catálogo</h2>
        <a class="section__link" href="<%=AppBase()%>artists.asp">Ver todos <%=Icon("chevron-right", "")%></a>
    </div>
    <div class="grid grid--artists">
        <% Do Until rsArtists.EOF %>
        <%=ArtistTile(rsArtists("ID_ARTIST"), Trim(rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME") & ""), rsArtists("ARTIST_IMAGE") & "", rsArtists("SONGS"))%>
        <% rsArtists.MoveNext : Loop %>
    </div>
</section>

<% rsSongs.Close : rsAlbums.Close : rsArtists.Close : rsPl.Close %>
<!--#include file="includes/layout_bottom.asp"-->
