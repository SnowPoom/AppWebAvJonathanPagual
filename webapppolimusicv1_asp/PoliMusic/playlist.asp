<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
RequireLogin()

Dim idPl, rsPl, rsSongs, rsAvailable, action, isOwner, canView
Dim errMsg, fTitle, fType, fPhoto, photoCsv, photoArr, i, idx, totalPlays
Dim idSong, songExists, alreadyIn, extra

idPl = ToInt(QS("id"), 0)
If idPl <= 0 Then
    Call Flash("Playlist no especificada.", "error")
    GoToPage AppBase() & "playlists.asp"
End If

Set rsPl = DbQuery("SELECT p.ID_PLAYLIST, p.ID_USER, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME " & _
    "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER WHERE p.ID_PLAYLIST = ?", Array(idPl))
If rsPl.EOF Then
    rsPl.Close
    Call Flash("La playlist solicitada no existe.", "error")
    GoToPage AppBase() & "playlists.asp"
End If

isOwner = (ToInt(rsPl("ID_USER"), 0) = CurrentUserId())
canView = isOwner Or IsAdmin() Or (ToInt(rsPl("TYPE"), 0) = 0)
If Not canView Then
    rsPl.Close
    Call Flash("Esta playlist es privada.", "error")
    GoToPage AppBase() & "playlists.asp"
End If

' ------------------------- Acciones POST -------------------------
action = FM("action")
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not CsrfValid() Then
        Call Flash("La sesión expiró. Vuelve a intentarlo.", "error")
        GoToPage AppBase() & "playlist.asp?id=" & idPl
    ElseIf Not isOwner Then
        Call Flash("Solo el propietario puede modificar esta playlist.", "error")
        GoToPage AppBase() & "playlist.asp?id=" & idPl
    ElseIf action = "update" Then
        fTitle = FM("title")
        fType = ToInt(FM("type"), 0)
        fPhoto = FM("photo")
        If fType <> 0 And fType <> 1 Then fType = 0
        If fPhoto = "" Then fPhoto = rsPl("PHOTO") & ""
        If fTitle = "" Then
            Call Flash("El título no puede quedar vacio.", "error")
        ElseIf Len(fTitle) > 50 Then
            Call Flash("El título no puede superar los 50 caracteres.", "error")
        Else
            Call DbExecute("UPDATE TBL_PLAYLIST SET TITLE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = ?", _
                Array(fTitle, fType, fPhoto, idPl))
            Call Flash("Playlist actualizada correctamente.", "success")
        End If
        GoToPage AppBase() & "playlist.asp?id=" & idPl

    ElseIf action = "delete" Then
        Call DbExecute("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?", Array(idPl))
        Call DbExecute("DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?", Array(idPl))
        Call Flash("La playlist se elimino junto con sus canciones asociadas.", "success")
        GoToPage AppBase() & "playlists.asp"

    ElseIf action = "remove_song" Then
        idSong = ToInt(FM("song"), 0)
        Call DbExecute("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(idPl, idSong))
        Call Flash("Canción quitada de la playlist.", "success")
        GoToPage AppBase() & "playlist.asp?id=" & idPl

    ElseIf action = "add_song" Then
        idSong = ToInt(FM("song"), 0)
        If idSong > 0 Then
            songExists = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong)), 0)
            alreadyIn = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?", Array(idPl, idSong)), 0)
            If songExists = 0 Then
                Call Flash("La canción seleccionada no existe.", "error")
            ElseIf alreadyIn > 0 Then
                Call Flash("Esa canción ya forma parte de la playlist.", "info")
            Else
                Call DbExecute("INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)", Array(idPl, idSong))
                Call Flash("Canción agregada a la playlist.", "success")
            End If
        End If
        GoToPage AppBase() & "playlist.asp?id=" & idPl
    End If
End If

pageTitle = rsPl("TITLE") & ""
pageSubtitle = "Detalle de la playlist"
navActive = "playlists"

Set rsSongs = DbQuery("SELECT * FROM (" & SongBaseSql() & _
    " WHERE EXISTS (SELECT 1 FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_SONG = s.ID_SONG AND ps.ID_PLAYLIST = ?)) q " & _
    "ORDER BY q.SONG_NAME", Array(idPl))

Set rsAvailable = DbQuery("SELECT s.ID_SONG, s.SONG_NAME FROM TBL_SONG s WHERE NOT EXISTS " & _
    "(SELECT 1 FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_SONG = s.ID_SONG AND ps.ID_PLAYLIST = ?) ORDER BY s.SONG_NAME", Array(idPl))

photoCsv = ListFilesCsv(DIR_PLAYLISTS, Array("jpg", "jpeg", "png", "webp", "gif"))
If photoCsv = "" Then photoCsv = "~/playlistImages/default.jpg"
photoArr = Split(photoCsv, "|")

totalPlays = 0
%>
<!--#include file="includes/layout_top.asp"-->

<a class="btn btn--ghost btn--sm" href="<%=AppBase()%>playlists.asp" style="margin-bottom:16px"><%=Icon("arrow-left", "")%> Volver a mis playlists</a>

<section class="detail">
    <div class="detail__media">
        <%=Cover(rsPl("PHOTO") & "", rsPl("TITLE") & "", "")%>
    </div>
    <div class="detail__body">
        <div class="detail__kind">Playlist</div>
        <h2><%=HtmlEncode(rsPl("TITLE") & "")%></h2>
        <div class="detail__meta">
            <span class="chip<%=IIf(ToInt(rsPl("TYPE"), 0) = 1, " chip--private", " chip--public")%>"><%=IIf(ToInt(rsPl("TYPE"), 0) = 1, "Privada", "Pública")%></span>
            <span class="dot"></span>
            <span><%=Icon("user", "")%>&#32;<%=HtmlEncode(rsPl("USERNAME") & "")%></span>
            <span class="dot"></span>
            <span><%=Icon("calendar", "")%>&#32;<%=FormatDateShort(rsPl("CREATION_DATE"))%></span>
            <span class="dot"></span>
            <span><%=rsSongs.RecordCount%>&#32;<%=IIf(rsSongs.RecordCount = 1, "canción", "canciones")%></span>
        </div>
        <div class="detail__actions">
            <button class="btn btn--primary" type="button" data-play-all><%=Icon("play", "")%> Reproducir todo</button>
            <% If isOwner Then %>
            <a class="btn btn--ghost" href="<%=AppBase()%>playlist.asp?id=<%=idPl%>&edit=1"><%=Icon("edit", "")%> Editar</a>
            <form method="post" action="<%=AppBase()%>playlist.asp?id=<%=idPl%>" style="display:inline">
                <%=CsrfField()%>
                <input type="hidden" name="action" value="delete">
                <button class="btn btn--danger" type="submit" data-confirm="Se eliminara la playlist y sus canciones asociadas. Esta acción no se puede deshacer." data-confirm-title="Eliminar playlist" data-confirm-ok="Si, eliminar">
                    <%=Icon("trash", "")%> Eliminar
                </button>
            </form>
            <% End If %>
        </div>
    </div>
</section>

<% If isOwner And QS("edit") = "1" Then %>
<section class="panel mt-24">
    <div class="panel__head">
        <h3><%=Icon("edit", "")%> Editar playlist</h3>
        <a class="icon-btn icon-btn--ghost" href="<%=AppBase()%>playlist.asp?id=<%=idPl%>" aria-label="Cerrar">&times;</a>
    </div>
    <form method="post" action="<%=AppBase()%>playlist.asp?id=<%=idPl%>">
        <%=CsrfField()%>
        <input type="hidden" name="action" value="update">
        <div class="form-grid">
            <div class="field">
                <label for="edTitle">Título</label>
                <input class="input" id="edTitle" name="title" maxlength="50" required value="<%=HtmlAttr(rsPl("TITLE") & "")%>">
            </div>
            <div class="field">
                <label for="edType">Visibilidad</label>
                <select class="select" id="edType" name="type">
                    <option value="0"<%=IIf(ToInt(rsPl("TYPE"), 0) = 0, " selected", "")%>>Pública</option>
                    <option value="1"<%=IIf(ToInt(rsPl("TYPE"), 0) = 1, " selected", "")%>>Privada</option>
                </select>
            </div>
            <div class="field field--full">
                <label>Portada</label>
                <div class="pick-grid">
                    <% For i = 0 To UBound(photoArr) %>
                    <input type="radio" id="edph<%=i%>" name="photo" value="<%=HtmlAttr(photoArr(i))%>"<%=IIf(photoArr(i) = (rsPl("PHOTO") & ""), " checked", "")%>>
                    <label for="edph<%=i%>"><%=Cover(photoArr(i), "Playlist " & (i + 1), "")%></label>
                    <% Next %>
                </div>
            </div>
        </div>
        <div class="form-actions">
            <button class="btn btn--primary" type="submit"><%=Icon("save", "")%> Guardar cambios</button>
            <a class="btn btn--ghost" href="<%=AppBase()%>playlist.asp?id=<%=idPl%>">Cancelar</a>
        </div>
    </form>
</section>
<% End If %>

<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("music", "")%> Canciones de la playlist</h2>
    </div>

    <% If rsSongs.EOF Then %>
        <%=EmptyState("music", "Playlist vacía", "Agrega canciones desde el catálogo o con el selector de abajo.", "<a class=""btn btn--ghost btn--sm"" href=""" & AppBase() & "songs.asp"">Ir al catálogo</a>")%>
    <% Else %>
    <div class="panel panel--flush">
        <div class="songlist">
            <%
            idx = 1
            Do Until rsSongs.EOF
                totalPlays = totalPlays + ToInt(rsSongs("PLAYS"), 0)
                extra = ""
                If isOwner Then
                    extra = "<form method=""post"" action=""" & AppBase() & "playlist.asp?id=" & idPl & """ style=""display:flex"">" & _
                            CsrfField() & "<input type=""hidden"" name=""action"" value=""remove_song"">" & _
                            "<input type=""hidden"" name=""song"" value=""" & rsSongs("ID_SONG") & """>" & _
                            "<button class=""icon-btn icon-btn--danger"" type=""submit"" title=""Quitar de la playlist"" aria-label=""Quitar"">" & _
                            Icon("minus", "") & "</button></form>"
                Else
                    extra = ""
                End If
                Response.Write SongRowHtml(rsSongs("ID_SONG"), rsSongs("SONG_NAME") & "", rsSongs("SONG_PATH") & "", rsSongs("PLAYS"), _
                    rsSongs("ARTISTS") & "", rsSongs("ALBUM_TITLE") & "", rsSongs("ALBUM_IMAGE") & "", extra, idx)
                idx = idx + 1
                rsSongs.MoveNext
            Loop
            %>
        </div>
    </div>
    <p class="small muted mt-8">Reproducciones acumuladas de la playlist: <b class="mono"><%=totalPlays%></b></p>
    <% End If %>
</section>

<% If isOwner Then %>
<section class="section">
    <div class="panel">
        <div class="panel__head">
            <h3><%=Icon("plus", "")%> Agregar canciones</h3>
        </div>
        <% If rsAvailable.EOF Then %>
            <p class="small muted">Todas las canciones del catálogo ya estan en esta playlist.</p>
        <% Else %>
        <form method="post" action="<%=AppBase()%>playlist.asp?id=<%=idPl%>" class="row row--wrap">
            <%=CsrfField()%>
            <input type="hidden" name="action" value="add_song">
            <select class="select" name="song" style="flex:1 1 260px" aria-label="Canción a agregar">
                <% Do Until rsAvailable.EOF %>
                <option value="<%=rsAvailable("ID_SONG")%>"><%=HtmlEncode(rsAvailable("SONG_NAME") & "")%></option>
                <% rsAvailable.MoveNext : Loop %>
            </select>
            <button class="btn btn--soft" type="submit"><%=Icon("plus", "")%> Agregar</button>
        </form>
        <% End If %>
    </div>
</section>
<% End If %>

<% rsPl.Close : rsSongs.Close : rsAvailable.Close %>
<!--#include file="includes/layout_bottom.asp"-->
