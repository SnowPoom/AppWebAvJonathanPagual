<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
RequireAdmin()

Dim idArtist, isEdit, rs, errMsg, conn
Dim fName, fLast, fImage, imgCsv, imgArr, i, imgCount, songCount

idArtist = ToInt(QS("id"), 0)
isEdit = (idArtist > 0)
errMsg = ""
fName = ""
fLast = ""
fImage = ""
songCount = 0

If isEdit Then
    Set rs = DbQuery("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?", Array(idArtist))
    If rs.EOF Then
        rs.Close
        Call Flash("El artista solicitado no existe.", "error")
        GoToPage AppBase() & "admin/artists.asp"
    End If
    fName = rs("ARTIST_NAME") & ""
    fLast = rs("ARTIST_LASTNAME") & ""
    fImage = rs("ARTIST_IMAGE") & ""
    songCount = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = ?", Array(idArtist)), 0)
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    fName = FM("artist_name")
    fLast = FM("artist_lastname")
    fImage = FM("artist_image")

    If fImage <> "" And InStr(fImage, "/") = 0 And InStr(fImage, "\") = 0 Then
        fImage = "~/artistImages/" & fImage
    End If

    If Not CsrfValid() Then
        errMsg = "La sesión expiró. Vuelve a intentarlo."
    ElseIf fName = "" Then
        errMsg = "El nombre del artista es obligatorio."
    ElseIf Len(fName) > 100 Or Len(fLast) > 100 Then
        errMsg = "El nombre y el apellido admiten hasta 100 caracteres."
    ElseIf Len(fImage) > 255 Then
        errMsg = "La ruta de la imagen no puede superar los 255 caracteres."
    Else
        If isEdit Then
            Call DbExecute("UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?", _
                Array(fName, fLast, fImage, idArtist))
            Call Flash("Artista '" & fName & "' actualizado.", "success")
        Else
            Call DbExecute("INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)", _
                Array(fName, fLast, fImage))
            Call Flash("Artista '" & fName & "' creado.", "success")
        End If
        GoToPage AppBase() & "admin/artists.asp"
    End If
End If

pageTitle = IIf(isEdit, "Editar artista", "Nuevo artista")
pageSubtitle = IIf(isEdit, "Actualiza los datos del artista", "Registra un artista en el catálogo")
navActive = "admin_artists"

imgCsv = ListFilesCsv(DIR_ARTISTS, Array("jpg", "jpeg", "png", "webp", "gif"))
imgArr = Split(imgCsv, "|")
imgCount = 0
If imgCsv <> "" Then imgCount = UBound(imgArr) + 1
%>
<!--#include file="../includes/layout_top.asp"-->

<a class="btn btn--ghost btn--sm" href="<%=AppBase()%>admin/artists.asp" style="margin-bottom:16px"><%=Icon("arrow-left", "")%> Volver a artistas</a>

<section class="panel">
    <div class="panel__head">
        <h3><%=Icon("mic", "")%>&#32;<%=HtmlEncode(pageTitle)%></h3>
        <% If isEdit Then %>
        <span class="chip chip--muted">ID <%=idArtist%> &middot; <%=songCount%> canciones</span>
        <% End If %>
    </div>

    <% If errMsg <> "" Then %>
    <div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
    <% End If %>

    <form method="post" action="<%=AppBase()%>admin/artist_form.asp<%=IIf(isEdit, "?id=" & idArtist, "")%>">
        <%=CsrfField()%>
        <div class="form-grid">
            <div class="field">
                <label for="aName">Nombre</label>
                <input class="input" id="aName" name="artist_name" maxlength="100" required value="<%=HtmlAttr(fName)%>">
            </div>
            <div class="field">
                <label for="aLast">Apellido</label>
                <input class="input" id="aLast" name="artist_lastname" maxlength="100" value="<%=HtmlAttr(fLast)%>">
            </div>
            <div class="field field--full">
                <label for="aImg">Ruta de la imagen</label>
                <input class="input" id="aImg" name="artist_image" maxlength="255" value="<%=HtmlAttr(fImage)%>"
                       placeholder="~/artistImages/artista.jpg" list="imgList">
                <datalist id="imgList">
                    <% For i = 0 To UBound(imgArr) %>
                    <% If imgArr(i) <> "" Then %><option value="<%=HtmlAttr(imgArr(i))%>"></option><% End If %>
                    <% Next %>
                </datalist>
                <span class="field__hint">
                    <%=imgCount%> imagen(es) en <b>artistImages</b>.
                    <% If fImage <> "" Then %>Estado: <b><%=IIf(AssetExists(fImage), "archivo encontrado", "archivo no encontrado")%></b>.<% End If %>
                    Si se deja vacia se usa un avatar generado con las iniciales.
                </span>
            </div>
            <% If imgCount > 0 Then %>
            <div class="field field--full">
                <label>Imágenes disponibles (clic para asignar la ruta)</label>
                <div class="pick-grid" data-pick-grid data-target="#aImg">
                    <% For i = 0 To UBound(imgArr) %>
                    <% If imgArr(i) <> "" Then %>
                    <button type="button" class="thumb-pick" data-path="<%=HtmlAttr(imgArr(i))%>" title="<%=HtmlAttr(imgArr(i))%>">
                        <%=Cover(imgArr(i), "Artista " & (i + 1), "cover--rounded")%>
                    </button>
                    <% End If %>
                    <% Next %>
                </div>
            </div>
            <% End If %>
        </div>
        <div class="form-actions">
            <button class="btn btn--primary" type="submit"><%=Icon("save", "")%>&#32;<%=IIf(isEdit, "Guardar cambios", "Crear artista")%></button>
            <a class="btn btn--ghost" href="<%=AppBase()%>admin/artists.asp">Cancelar</a>
        </div>
    </form>
</section>

<% If isEdit Then rs.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
