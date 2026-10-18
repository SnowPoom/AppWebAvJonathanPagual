<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Panel de administración"
pageSubtitle = "Resumen general de PoliMusic"
navActive = "admin"
RequireAdmin()

Dim nUsers, nAdmins, nSongs, nArtists, nAlbums, nPlaylists, nLinks
Dim rsTop, rsUsers, rsMissing, idx, missingCount, s

nUsers = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER", Empty), 0)
nAdmins = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1", Empty), 0)
nSongs = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG", Empty), 0)
nArtists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ARTIST", Empty), 0)
nAlbums = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_ALBUM", Empty), 0)
nPlaylists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST", Empty), 0)
nLinks = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG", Empty), 0)

Set rsTop = DbQuery("SELECT TOP 5 * FROM (" & SongBaseSql() & ") q ORDER BY q.PLAYS DESC, q.ID_SONG", Empty)
Set rsUsers = DbQuery("SELECT TOP 5 u.ID_USER, u.USERNAME, u.EMAIL, u.USER_TYPE, u.USER_PHOTO, u.BIRTHDAY, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS PLAYLISTS " & _
    "FROM TBL_USER u ORDER BY u.ID_USER DESC", Empty)
Set rsMissing = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH FROM TBL_SONG ORDER BY SONG_NAME", Empty)
%>
<!--#include file="../includes/layout_top.asp"-->

<div class="grid grid--stats">
    <a class="stat" href="<%=AppBase()%>admin/users.asp">
        <span class="stat__ico"><%=Icon("users", "")%></span>
        <div><b><%=nUsers%></b><span>Usuarios (<%=nAdmins%> admin)</span></div>
    </a>
    <a class="stat" href="<%=AppBase()%>admin/songs.asp">
        <span class="stat__ico"><%=Icon("music", "")%></span>
        <div><b><%=nSongs%></b><span>Canciones</span></div>
    </a>
    <a class="stat" href="<%=AppBase()%>admin/artists.asp">
        <span class="stat__ico"><%=Icon("mic", "")%></span>
        <div><b><%=nArtists%></b><span>Artistas</span></div>
    </a>
    <a class="stat" href="<%=AppBase()%>admin/albums.asp">
        <span class="stat__ico"><%=Icon("disc", "")%></span>
        <div><b><%=nAlbums%></b><span>Álbumes</span></div>
    </a>
    <div class="stat">
        <span class="stat__ico"><%=Icon("list", "")%></span>
        <div><b><%=nPlaylists%></b><span>Playlists (<%=nLinks%> pistas)</span></div>
    </div>
</div>

<div class="grid grid--two mt-24">
    <section class="panel">
        <div class="panel__head">
            <h3><%=Icon("wave", "")%> Canciones más escuchadas</h3>
            <a class="section__link" href="<%=AppBase()%>admin/songs.asp">Gestionar <%=Icon("chevron-right", "")%></a>
        </div>
        <div class="table-wrap" style="border:0;background:none">
            <table class="data" style="min-width:0">
                <thead>
                    <tr><th>#</th><th>Canción</th><th>Artistas</th><th class="num">Reproducciones</th></tr>
                </thead>
                <tbody>
                    <% idx = 1 %>
                    <% Do Until rsTop.EOF %>
                    <tr>
                        <td class="num muted"><%=idx%></td>
                        <td><strong><%=HtmlEncode(rsTop("SONG_NAME") & "")%></strong></td>
                        <td><%=HtmlEncode(TextOr(rsTop("ARTISTS") & "", "-"))%></td>
                        <td class="num"><%=ToInt(rsTop("PLAYS"), 0)%></td>
                    </tr>
                    <% idx = idx + 1 %>
                    <% rsTop.MoveNext : Loop %>
                </tbody>
            </table>
        </div>
    </section>

    <section class="panel">
        <div class="panel__head">
            <h3><%=Icon("users", "")%> Usuarios recientes</h3>
            <a class="section__link" href="<%=AppBase()%>admin/users.asp">Ver todos <%=Icon("chevron-right", "")%></a>
        </div>
        <div class="stack" style="gap:6px">
            <% Do Until rsUsers.EOF %>
            <div class="row" style="padding:8px 10px;border-radius:12px;background:rgba(125,211,252,.04)">
                <%=Avatar(rsUsers("USER_PHOTO") & "", rsUsers("USERNAME") & "", "avatar--sm")%>
                <div style="min-width:0;flex:1">
                    <div class="truncate" style="font-size:13.5px;font-weight:700"><%=HtmlEncode(rsUsers("USERNAME") & "")%></div>
                    <div class="truncate small muted"><%=HtmlEncode(rsUsers("EMAIL") & "")%></div>
                </div>
                <span class="chip<%=IIf(ToInt(rsUsers("USER_TYPE"), 0) = 1, " chip--admin", "")%>"><%=RoleLabel(rsUsers("USER_TYPE"))%></span>
            </div>
            <% rsUsers.MoveNext : Loop %>
        </div>
    </section>
</div>

<section class="section">
    <div class="panel">
        <div class="panel__head">
            <h3><%=Icon("alert", "")%> Estado de los archivos de audio</h3>
            <a class="section__link" href="<%=AppBase()%>admin/songs.asp">Revisar canciones <%=Icon("chevron-right", "")%></a>
        </div>
        <p class="small muted" style="margin-bottom:14px">
            PoliMusic verifica en disco cada ruta guardada en <b>TBL_SONG.SONG_PATH</b>. Los archivos que no existen se marcan como <b>Sin audio</b> en el catálogo.
        </p>
        <div class="table-wrap">
            <table class="data">
                <thead><tr><th>Canción</th><th>Ruta registrada</th><th>Estado</th></tr></thead>
                <tbody>
                    <% missingCount = 0 %>
                    <% Do Until rsMissing.EOF %>
                    <% If Not AssetExists(rsMissing("SONG_PATH") & "") Then missingCount = missingCount + 1 %>
                    <tr>
                        <td><strong><%=HtmlEncode(rsMissing("SONG_NAME") & "")%></strong></td>
                        <td class="small mono"><%=HtmlEncode(rsMissing("SONG_PATH") & "")%></td>
                        <td>
                            <% If AssetExists(rsMissing("SONG_PATH") & "") Then %>
                            <span class="chip chip--admin">Disponible</span>
                            <% Else %>
                            <span class="chip chip--muted">Falta el archivo</span>
                            <% End If %>
                        </td>
                    </tr>
                    <% rsMissing.MoveNext : Loop %>
                </tbody>
            </table>
        </div>
        <% If missingCount > 0 Then %>
        <p class="small muted mt-8">Hay <b class="mono"><%=missingCount%></b> archivo(s) pendiente(s). Copia los MP3 en la carpeta <b>songFiles</b> del sitio o edita la ruta desde el modulo de canciones.</p>
        <% End If %>
    </div>
</section>

<% rsTop.Close : rsUsers.Close : rsMissing.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
