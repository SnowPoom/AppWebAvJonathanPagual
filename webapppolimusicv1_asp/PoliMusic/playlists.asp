<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Mis playlists"
pageSubtitle = "Organiza tu música en colecciones propias"
navActive = "playlists"
RequireLogin()

Dim showForm, errMsg, fTitle, fType, fPhoto, action
Dim rsMine, rsPublic, photoCsv, photoArr, i

showForm = (QS("nueva") = "1")
errMsg = ""
fTitle = ""
fType = 0
fPhoto = "~/playlistImages/default.jpg"
action = FM("action")

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not CsrfValid() Then
        errMsg = "La sesión expiró. Vuelve a intentarlo."
        showForm = True
    ElseIf action = "create" Then
        showForm = True
        fTitle = FM("title")
        fType = ToInt(FM("type"), 0)
        fPhoto = FM("photo")
        If fType <> 0 And fType <> 1 Then fType = 0
        If fPhoto = "" Then fPhoto = "~/playlistImages/default.jpg"

        If fTitle = "" Then
            errMsg = "El título de la playlist es obligatorio."
        ElseIf Len(fTitle) > 50 Then
            errMsg = "El título no puede superar los 50 caracteres."
        ElseIf ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ? AND TITLE = ?", Array(CurrentUserId(), fTitle)), 0) > 0 Then
            errMsg = "Ya tienes una playlist con ese título."
        Else
            Call DbExecute("INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (?, ?, GETDATE(), ?, ?)", _
                Array(CurrentUserId(), fTitle, fType, fPhoto))
            Call Flash("Playlist '" & fTitle & "' creada correctamente.", "success")
            GoToPage AppBase() & "playlists.asp"
        End If
    End If
End If

Set rsMine = DbQuery("SELECT p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, p.CREATION_DATE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONGS " & _
    "FROM TBL_PLAYLIST p WHERE p.ID_USER = ? ORDER BY p.CREATION_DATE DESC, p.ID_PLAYLIST DESC", Array(CurrentUserId()))

Set rsPublic = DbQuery("SELECT TOP 8 p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, p.CREATION_DATE, u.USERNAME, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS SONGS " & _
    "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER " & _
    "WHERE p.TYPE = 0 AND p.ID_USER <> ? ORDER BY p.CREATION_DATE DESC", Array(CurrentUserId()))

photoCsv = ListFilesCsv(DIR_PLAYLISTS, Array("jpg", "jpeg", "png", "webp", "gif"))
If photoCsv = "" Then photoCsv = "~/playlistImages/default.jpg"
photoArr = Split(photoCsv, "|")
%>
<!--#include file="includes/layout_top.asp"-->

<div class="admin-head">
    <div>
        <h2><%=rsMine.RecordCount%>&#32;<%=IIf(rsMine.RecordCount = 1, "playlist propia", "playlists propias")%></h2>
        <p>Las playlists privadas solo son visibles para ti.</p>
    </div>
    <a class="btn btn--primary" href="<%=AppBase()%>playlists.asp?nueva=1"><%=Icon("plus", "")%> Nueva playlist</a>
</div>

<% If showForm Then %>
<section class="panel" style="margin-bottom:22px">
    <div class="panel__head">
        <h3><%=Icon("list", "")%> Crear nueva playlist</h3>
        <a class="icon-btn icon-btn--ghost" href="<%=AppBase()%>playlists.asp" aria-label="Cerrar formulario">&times;</a>
    </div>

    <% If errMsg <> "" Then %>
    <div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
    <% End If %>

    <form method="post" action="<%=AppBase()%>playlists.asp">
        <%=CsrfField()%>
        <input type="hidden" name="action" value="create">
        <div class="form-grid">
            <div class="field">
                <label for="plTitle">Título</label>
                <input class="input" id="plTitle" name="title" maxlength="50" required value="<%=HtmlAttr(fTitle)%>"
                       placeholder="Ej. Mis favoritas de siempre">
                <span class="field__hint">Máximo 50 caracteres.</span>
            </div>
            <div class="field">
                <label for="plType">Visibilidad</label>
                <select class="select" id="plType" name="type">
                    <option value="0"<%=IIf(fType = 0, " selected", "")%>>Pública - visible para todos</option>
                    <option value="1"<%=IIf(fType = 1, " selected", "")%>>Privada - solo para mi</option>
                </select>
                <span class="field__hint">El tipo se guarda en TBL_PLAYLIST.TYPE (0 = Pública, 1 = Privada).</span>
            </div>
            <div class="field field--full">
                <label>Portada</label>
                <div class="pick-grid">
                    <% For i = 0 To UBound(photoArr) %>
                    <input type="radio" id="ph<%=i%>" name="photo" value="<%=HtmlAttr(photoArr(i))%>"<%=IIf(photoArr(i) = fPhoto Or (fPhoto = "" And i = 0), " checked", "")%>>
                    <label for="ph<%=i%>" title="<%=HtmlAttr(photoArr(i))%>">
                        <%=Cover(photoArr(i), "Playlist " & (i + 1), "")%>
                    </label>
                    <% Next %>
                </div>
            </div>
        </div>
        <div class="form-actions">
            <button class="btn btn--primary" type="submit"><%=Icon("save", "")%> Crear playlist</button>
            <a class="btn btn--ghost" href="<%=AppBase()%>playlists.asp">Cancelar</a>
        </div>
    </form>
</section>
<% End If %>

<% If rsMine.EOF Then %>
    <%=EmptyState("list", "Aun no tienes playlists", "Crea tu primera playlist y empieza a agrupar tus canciones favoritas.", "<a class=""btn btn--primary btn--sm"" href=""" & AppBase() & "playlists.asp?nueva=1"">Crear playlist</a>")%>
<% Else %>
<div class="grid grid--playlists">
    <% Do Until rsMine.EOF %>
    <%=PlaylistTile(rsMine("ID_PLAYLIST"), rsMine("TITLE") & "", rsMine("TYPE"), rsMine("PHOTO") & "", rsMine("SONGS"), rsMine("CREATION_DATE"), Session("Usuario") & "")%>
    <% rsMine.MoveNext : Loop %>
</div>
<% End If %>

<% If Not rsPublic.EOF Then %>
<section class="section">
    <div class="section__head">
        <h2 class="section__title"><%=Icon("sparkles", "")%> Playlists públicas de la comunidad</h2>
    </div>
    <div class="grid grid--playlists">
        <% Do Until rsPublic.EOF %>
        <%=PlaylistTile(rsPublic("ID_PLAYLIST"), rsPublic("TITLE") & "", rsPublic("TYPE"), rsPublic("PHOTO") & "", rsPublic("SONGS"), rsPublic("CREATION_DATE"), rsPublic("USERNAME") & "")%>
        <% rsPublic.MoveNext : Loop %>
    </div>
</section>
<% End If %>

<% rsMine.Close : rsPublic.Close %>
<!--#include file="includes/layout_bottom.asp"-->
