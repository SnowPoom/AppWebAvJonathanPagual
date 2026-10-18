<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
RequireAdmin()

Dim idSong, isEdit, rs, errMsg, action, conn
Dim fName, fPath, fPlays, fArtists, fAlbums
Dim rsArtists, rsAlbums, filesCsv, filesArr, i, newId, csv, arr, k, idItem, fileCount

idSong = ToInt(QS("id"), 0)
isEdit = (idSong > 0)
errMsg = ""
fName = ""
fPath = ""
fPlays = 0
fArtists = ""
fAlbums = ""

If isEdit Then
    Set rs = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH, ISNULL(PLAYS,0) AS PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array(idSong))
    If rs.EOF Then
        rs.Close
        Call Flash("La canción solicitada no existe.", "error")
        GoToPage AppBase() & "admin/songs.asp"
    End If
    fName = rs("SONG_NAME") & ""
    fPath = rs("SONG_PATH") & ""
    fPlays = ToInt(rs("PLAYS"), 0)

    Set rsArtists = DbQuery("SELECT ID_ARTIST FROM TBL_PLAYER WHERE ID_SONG = ?", Array(idSong))
    Do Until rsArtists.EOF
        If fArtists <> "" Then fArtists = fArtists & ","
        fArtists = fArtists & rsArtists("ID_ARTIST")
        rsArtists.MoveNext
    Loop
    rsArtists.Close

    Set rsAlbums = DbQuery("SELECT ID_ALBUM FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(idSong))
    Do Until rsAlbums.EOF
        If fAlbums <> "" Then fAlbums = fAlbums & ","
        fAlbums = fAlbums & rsAlbums("ID_ALBUM")
        rsAlbums.MoveNext
    Loop
    rsAlbums.Close
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    fName = FM("song_name")
    fPath = FM("song_path")
    fPlays = ToInt(FM("plays"), 0)
    fArtists = Request.Form("artists") & ""
    fAlbums = Request.Form("albums") & ""

    If fPath <> "" And InStr(fPath, "/") = 0 And InStr(fPath, "\") = 0 Then
        fPath = "~/songFiles/" & fPath
    End If
    If fPlays < 0 Then fPlays = 0

    If Not CsrfValid() Then
        errMsg = "La sesión expiró. Vuelve a intentarlo."
    ElseIf fName = "" Then
        errMsg = "El nombre de la canción es obligatorio."
    ElseIf Len(fName) > 50 Then
        errMsg = "El nombre no puede superar los 50 caracteres."
    ElseIf fPath = "" Then
        errMsg = "Debes indicar la ruta del archivo de audio."
    ElseIf Len(fPath) > 255 Then
        errMsg = "La ruta no puede superar los 255 caracteres."
    Else
        Set conn = Db()
        On Error Resume Next
        conn.BeginTrans

        If isEdit Then
            Call DbExecute("UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?", _
                Array(fName, fPath, fPlays, idSong))
            newId = idSong
        Else
            newId = ToInt(DbScalar("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) OUTPUT INSERTED.ID_SONG VALUES (?, ?, ?)", _
                Array(fName, fPath, fPlays)), 0)
        End If

        Call DbExecute("DELETE FROM TBL_PLAYER WHERE ID_SONG = ?", Array(newId))
        If fArtists <> "" Then
            arr = Split(fArtists, ",")
            For k = 0 To UBound(arr)
                idItem = ToInt(arr(k), 0)
                If idItem > 0 Then
                    Call DbExecute("INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)", Array(newId, idItem))
                End If
            Next
        End If

        Call DbExecute("DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?", Array(newId))
        If fAlbums <> "" Then
            arr = Split(fAlbums, ",")
            For k = 0 To UBound(arr)
                idItem = ToInt(arr(k), 0)
                If idItem > 0 Then
                    Call DbExecute("INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array(idItem, newId))
                End If
            Next
        End If

        If Err.Number <> 0 Then
            conn.RollbackTrans
            errMsg = "No se pudo guardar la canción: " & Err.Description
            Err.Clear
        Else
            conn.CommitTrans
            Call Flash("Canción '" & fName & "' guardada correctamente.", "success")
            GoToPage AppBase() & "admin/songs.asp"
        End If
        On Error GoTo 0
    End If
End If

pageTitle = IIf(isEdit, "Editar canción", "Nueva canción")
pageSubtitle = IIf(isEdit, "Modifica los datos y sus relaciones", "Registra una canción en el catálogo")
navActive = "admin_songs"

Set rsArtists = DbQuery("SELECT ID_ARTIST, LTRIM(RTRIM(ARTIST_NAME + ' ' + ISNULL(ARTIST_LASTNAME,''))) AS FULLNAME FROM TBL_ARTIST ORDER BY FULLNAME", Empty)
Set rsAlbums = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM FROM TBL_ALBUM ORDER BY TITLE_ALBUM", Empty)

filesCsv = ListFilesCsv(DIR_SONGS, Array("mp3", "wav", "ogg", "m4a", "aac"))
filesArr = Split(filesCsv, "|")
fileCount = 0
If filesCsv <> "" Then fileCount = UBound(filesArr) + 1
%>
<!--#include file="../includes/layout_top.asp"-->

<a class="btn btn--ghost btn--sm" href="<%=AppBase()%>admin/songs.asp" style="margin-bottom:16px"><%=Icon("arrow-left", "")%> Volver a canciones</a>

<section class="panel">
    <div class="panel__head">
        <h3><%=Icon("music", "")%>&#32;<%=HtmlEncode(pageTitle)%></h3>
        <% If isEdit Then %>
        <span class="chip chip--muted">ID <%=idSong%></span>
        <% End If %>
    </div>

    <% If errMsg <> "" Then %>
    <div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
    <% End If %>

    <form method="post" action="<%=AppBase()%>admin/song_form.asp<%=IIf(isEdit, "?id=" & idSong, "")%>">
        <%=CsrfField()%>
        <div class="form-grid">
            <div class="field">
                <label for="sName">Nombre de la canción</label>
                <input class="input" id="sName" name="song_name" maxlength="50" required value="<%=HtmlAttr(fName)%>">
            </div>
            <div class="field">
                <label for="sPlays">Reproducciones</label>
                <input class="input" id="sPlays" name="plays" type="number" min="0" step="1" value="<%=fPlays%>">
                <span class="field__hint">Se incrementa automáticamente al reproducir.</span>
            </div>
            <div class="field field--full">
                <label for="sPath">Ruta del archivo</label>
                <input class="input" id="sPath" name="song_path" maxlength="255" required list="fileList"
                       value="<%=HtmlAttr(fPath)%>" placeholder="~/songFiles/mi-canción.mp3">
                <datalist id="fileList">
                    <% For i = 0 To UBound(filesArr) %>
                    <% If filesArr(i) <> "" Then %>
                    <option value="<%=HtmlAttr(filesArr(i))%>"></option>
                    <% End If %>
                    <% Next %>
                </datalist>
                <span class="field__hint">
                    Archivos detectados en <b>songFiles</b>: <%=IIf(fileCount = 0, "ninguno", fileCount)%>.
                    <% If isEdit Then %>
                    Estado actual: <b><%=IIf(AssetExists(fPath), "archivo encontrado", "archivo no encontrado")%></b>.
                    <% End If %>
                </span>
            </div>

            <div class="field">
                <label>Artistas (TBL_PLAYER)</label>
                <div class="stack" style="gap:6px;max-height:220px;overflow:auto;padding:10px;border-radius:12px;background:rgba(5,20,33,.4);border:1px solid var(--glass-brd-soft)">
                    <% If rsArtists.EOF Then %>
                    <span class="small muted">No hay artistas registrados.</span>
                    <% End If %>
                    <% Do Until rsArtists.EOF %>
                    <% csv = "," & fArtists & "," %>
                    <label class="check">
                        <input type="checkbox" name="artists" value="<%=rsArtists("ID_ARTIST")%>"<%=IIf(InStr(csv, "," & rsArtists("ID_ARTIST") & ",") > 0, " checked", "")%>>
                        <%=HtmlEncode(rsArtists("FULLNAME") & "")%>
                    </label>
                    <% rsArtists.MoveNext : Loop %>
                </div>
            </div>

            <div class="field">
                <label>Álbumes (TBL_TRACKLIST)</label>
                <div class="stack" style="gap:6px;max-height:220px;overflow:auto;padding:10px;border-radius:12px;background:rgba(5,20,33,.4);border:1px solid var(--glass-brd-soft)">
                    <% If rsAlbums.EOF Then %>
                    <span class="small muted">No hay álbumes registrados.</span>
                    <% End If %>
                    <% Do Until rsAlbums.EOF %>
                    <% csv = "," & fAlbums & "," %>
                    <label class="check">
                        <input type="checkbox" name="albums" value="<%=rsAlbums("ID_ALBUM")%>"<%=IIf(InStr(csv, "," & rsAlbums("ID_ALBUM") & ",") > 0, " checked", "")%>>
                        <%=HtmlEncode(rsAlbums("TITLE_ALBUM") & "")%>
                    </label>
                    <% rsAlbums.MoveNext : Loop %>
                </div>
            </div>
        </div>
        <div class="form-actions">
            <button class="btn btn--primary" type="submit"><%=Icon("save", "")%>&#32;<%=IIf(isEdit, "Guardar cambios", "Crear canción")%></button>
            <a class="btn btn--ghost" href="<%=AppBase()%>admin/songs.asp">Cancelar</a>
        </div>
    </form>
</section>

<% If isEdit Then rs.Close %>
<% rsArtists.Close : rsAlbums.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
