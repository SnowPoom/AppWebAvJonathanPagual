<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
RequireAdmin()

Dim idAlbum, isEdit, rs, errMsg, conn, newId
Dim fTitle, fYear, fImage, fTracks
Dim rsSongs, imgCsv, imgArr, i, imgCount, csv, arr, k, idItem, trackCount

idAlbum = ToInt(QS("id"), 0)
isEdit = (idAlbum > 0)
errMsg = ""
fTitle = ""
fYear = ""
fImage = ""
fTracks = ""
trackCount = 0

If isEdit Then
    Set rs = DbQuery("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?", Array(idAlbum))
    If rs.EOF Then
        rs.Close
        Call Flash("El álbum solicitado no existe.", "error")
        GoToPage AppBase() & "admin/albums.asp"
    End If
    fTitle = rs("TITLE_ALBUM") & ""
    If Not IsNull(rs("RELEASE_YEAR")) Then fYear = rs("RELEASE_YEAR")
    fImage = rs("ALBUM_IMAGE_PATH") & ""

    Set rsSongs = DbQuery("SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(idAlbum))
    Do Until rsSongs.EOF
        If fTracks <> "" Then fTracks = fTracks & ","
        fTracks = fTracks & rsSongs("ID_SONG")
        rsSongs.MoveNext
    Loop
    rsSongs.Close
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    fTitle = FM("title_album")
    fYear = FM("release_year")
    fImage = FM("album_image")
    fTracks = Request.Form("tracks") & ""

    If fImage <> "" And InStr(fImage, "/") = 0 And InStr(fImage, "\") = 0 Then
        fImage = "~/albumImages/" & fImage
    End If

    If Not CsrfValid() Then
        errMsg = "La sesión expiró. Vuelve a intentarlo."
    ElseIf fTitle = "" Then
        errMsg = "El título del álbum es obligatorio."
    ElseIf Len(fTitle) > 50 Then
        errMsg = "El título no puede superar los 50 caracteres."
    ElseIf fYear <> "" And (ToInt(fYear, 0) < 1000 Or ToInt(fYear, 0) > 2100) Then
        errMsg = "El año de lanzamiento debe estar entre 1000 y 2100."
    ElseIf Len(fImage) > 255 Then
        errMsg = "La ruta de la portada no puede superar los 255 caracteres."
    Else
        Set conn = Db()
        On Error Resume Next
        conn.BeginTrans

        If isEdit Then
            Call DbExecute("UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?", _
                Array(fTitle, IIf(fYear = "", Null, ToInt(fYear, 0)), fImage, idAlbum))
            newId = idAlbum
        Else
            newId = ToInt(DbScalar("INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) OUTPUT INSERTED.ID_ALBUM " & _
                "VALUES (?, ?, ?)", Array(fTitle, IIf(fYear = "", Null, ToInt(fYear, 0)), fImage)), 0)
        End If

        Call DbExecute("DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?", Array(newId))
        If fTracks <> "" Then
            arr = Split(fTracks, ",")
            For k = 0 To UBound(arr)
                idItem = ToInt(arr(k), 0)
                If idItem > 0 Then
                    Call DbExecute("INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)", Array(newId, idItem))
                End If
            Next
        End If

        If Err.Number <> 0 Then
            conn.RollbackTrans
            errMsg = "No se pudo guardar el álbum: " & Err.Description
            Err.Clear
        Else
            conn.CommitTrans
            Call Flash("Álbum '" & fTitle & "' guardado correctamente.", "success")
            GoToPage AppBase() & "admin/albums.asp"
        End If
        On Error GoTo 0
    End If
End If

pageTitle = IIf(isEdit, "Editar álbum", "Nuevo álbum")
pageSubtitle = IIf(isEdit, "Actualiza los datos y su lista de canciones", "Registra un álbum en el catálogo")
navActive = "admin_albums"

Set rsSongs = DbQuery("SELECT s.ID_SONG, s.SONG_NAME, ISNULL((SELECT STRING_AGG(LTRIM(RTRIM(a.ARTIST_NAME)), ', ') " & _
    "FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG), '') AS ARTISTS " & _
    "FROM TBL_SONG s ORDER BY s.SONG_NAME", Empty)

imgCsv = ListFilesCsv(DIR_ALBUMS, Array("jpg", "jpeg", "png", "webp", "gif"))
imgArr = Split(imgCsv, "|")
imgCount = 0
If imgCsv <> "" Then imgCount = UBound(imgArr) + 1
%>
<!--#include file="../includes/layout_top.asp"-->

<a class="btn btn--ghost btn--sm" href="<%=AppBase()%>admin/albums.asp" style="margin-bottom:16px"><%=Icon("arrow-left", "")%> Volver a álbumes</a>

<section class="panel">
    <div class="panel__head">
        <h3><%=Icon("disc", "")%>&#32;<%=HtmlEncode(pageTitle)%></h3>
        <% If isEdit Then %><span class="chip chip--muted">ID <%=idAlbum%></span><% End If %>
    </div>

    <% If errMsg <> "" Then %>
    <div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
    <% End If %>

    <form method="post" action="<%=AppBase()%>admin/album_form.asp<%=IIf(isEdit, "?id=" & idAlbum, "")%>">
        <%=CsrfField()%>
        <div class="form-grid">
            <div class="field">
                <label for="alTitle">Título del álbum</label>
                <input class="input" id="alTitle" name="title_album" maxlength="50" required value="<%=HtmlAttr(fTitle)%>">
            </div>
            <div class="field">
                <label for="alYear">Año de lanzamiento</label>
                <input class="input" id="alYear" name="release_year" type="number" min="1000" max="2100" value="<%=HtmlAttr(fYear)%>">
            </div>
            <div class="field field--full">
                <label for="alImg">Ruta de la portada</label>
                <input class="input" id="alImg" name="album_image" maxlength="255" value="<%=HtmlAttr(fImage)%>"
                       placeholder="~/albumImages/album1.png" list="imgList">
                <datalist id="imgList">
                    <% For i = 0 To UBound(imgArr) %>
                    <% If imgArr(i) <> "" Then %><option value="<%=HtmlAttr(imgArr(i))%>"></option><% End If %>
                    <% Next %>
                </datalist>
                <span class="field__hint">
                    <%=imgCount%> imagen(es) en <b>albumImages</b>.
                    <% If fImage <> "" Then %>Estado: <b><%=IIf(AssetExists(fImage), "archivo encontrado", "archivo no encontrado")%></b>.<% End If %>
                    Si se deja vacia se usa una portada generada con las iniciales.
                </span>
            </div>
            <% If imgCount > 0 Then %>
            <div class="field field--full">
                <label>Portadas disponibles (clic para asignar la ruta)</label>
                <div class="pick-grid" data-pick-grid data-target="#alImg">
                    <% For i = 0 To UBound(imgArr) %>
                    <% If imgArr(i) <> "" Then %>
                    <button type="button" class="thumb-pick" data-path="<%=HtmlAttr(imgArr(i))%>" title="<%=HtmlAttr(imgArr(i))%>">
                        <%=Cover(imgArr(i), "Álbum " & (i + 1), "")%>
                    </button>
                    <% End If %>
                    <% Next %>
                </div>
            </div>
            <% End If %>
            <div class="field field--full">
                <label>Lista de canciones (TBL_TRACKLIST)</label>
                <div class="stack" style="gap:6px;max-height:280px;overflow:auto;padding:12px;border-radius:12px;background:rgba(5,20,33,.4);border:1px solid var(--glass-brd-soft)">
                    <% If rsSongs.EOF Then %>
                    <span class="small muted">No hay canciones registradas en el catálogo.</span>
                    <% End If %>
                    <% Do Until rsSongs.EOF %>
                    <% csv = "," & fTracks & "," %>
                    <label class="check">
                        <input type="checkbox" name="tracks" value="<%=rsSongs("ID_SONG")%>"<%=IIf(InStr(csv, "," & rsSongs("ID_SONG") & ",") > 0, " checked", "")%>>
                        <%=HtmlEncode(rsSongs("SONG_NAME") & "")%>
                        <span class="small muted">- <%=HtmlEncode(TextOr(rsSongs("ARTISTS") & "", "sin artista"))%></span>
                    </label>
                    <% rsSongs.MoveNext : Loop %>
                </div>
            </div>
        </div>
        <div class="form-actions">
            <button class="btn btn--primary" type="submit"><%=Icon("save", "")%>&#32;<%=IIf(isEdit, "Guardar cambios", "Crear álbum")%></button>
            <a class="btn btn--ghost" href="<%=AppBase()%>admin/albums.asp">Cancelar</a>
        </div>
    </form>
</section>

<% If isEdit Then rs.Close %>
<% rsSongs.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
