<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
RequireLogin()

Dim idAlbum, rsAlbum, rsTracks, idx, totalPlays, title, year

idAlbum = ToInt(QS("id"), 0)
If idAlbum <= 0 Then
    Flash "Álbum no especificado.", "error"
    GoToPage AppBase() & "albums.asp"
End If

Set rsAlbum = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(idAlbum))
If rsAlbum.EOF Then
    rsAlbum.Close
    Flash "El álbum solicitado no existe.", "error"
    GoToPage AppBase() & "albums.asp"
End If

title = rsAlbum("TITLE_ALBUM") & ""
year = ToInt(rsAlbum("RELEASE_YEAR"), 0)
pageTitle = title
pageSubtitle = "Detalle del álbum"
navActive = "albums"

Set rsTracks = DbQuery("SELECT * FROM (" & SongBaseSql() & _
    " WHERE EXISTS (SELECT 1 FROM TBL_TRACKLIST t WHERE t.ID_SONG = s.ID_SONG AND t.ID_ALBUM = ?)) q " & _
    "ORDER BY q.ID_SONG", Array(idAlbum))

totalPlays = 0
%>
<!--#include file="includes/layout_top.asp"-->

<a class="btn btn--ghost btn--sm" href="<%=AppBase()%>albums.asp" style="margin-bottom:16px"><%=Icon("arrow-left", "")%> Volver a álbumes</a>

<section class="detail">
    <div class="detail__media">
        <%=Cover(rsAlbum("ALBUM_IMAGE_PATH") & "", title, "")%>
    </div>
    <div class="detail__body">
        <div class="detail__kind">Álbum</div>
        <h2><%=HtmlEncode(title)%></h2>
        <div class="detail__meta">
            <span class="chip"><%=Icon("calendar", "")%>&#32;<%=IIf(year > 0, year, "Sin año")%></span>
            <span class="dot"></span>
            <span><%=rsTracks.RecordCount%>&#32;<%=IIf(rsTracks.RecordCount = 1, "canción", "canciones")%></span>
        </div>
        <div class="detail__actions">
            <button class="btn btn--primary" type="button" data-play-all><%=Icon("play", "")%> Reproducir álbum</button>
            <a class="btn btn--ghost" href="<%=AppBase()%>songs.asp?álbum=<%=idAlbum%>"><%=Icon("music", "")%> Ver en el catálogo</a>
        </div>
    </div>
</section>

<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("list", "")%> Lista de canciones</h2>
        <span class="small muted"><%=rsTracks.RecordCount%> pistas</span>
    </div>
    <% If rsTracks.EOF Then %>
        <%=EmptyState("music", "Álbum sin canciones", "Este álbum aún no tiene pistas asociadas en TBL_TRACKLIST.", "")%>
    <% Else %>
    <div class="panel panel--flush">
        <div class="songlist">
            <%
            idx = 1
            Do Until rsTracks.EOF
                totalPlays = totalPlays + ToInt(rsTracks("PLAYS"), 0)
                Response.Write SongRowHtml(rsTracks("ID_SONG"), rsTracks("SONG_NAME") & "", rsTracks("SONG_PATH") & "", rsTracks("PLAYS"), _
                    rsTracks("ARTISTS") & "", rsTracks("ALBUM_TITLE") & "", rsTracks("ALBUM_IMAGE") & "", "", idx)
                idx = idx + 1
                rsTracks.MoveNext
            Loop
            %>
        </div>
    </div>
    <p class="small muted mt-8">Reproducciones totales del álbum: <b class="mono"><%=totalPlays%></b></p>
    <% End If %>
</section>

<% rsAlbum.Close : rsTracks.Close %>
<!--#include file="includes/layout_bottom.asp"-->
