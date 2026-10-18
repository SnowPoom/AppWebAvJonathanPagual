<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Mi perfil"
pageSubtitle = "Administra tus datos y tu seguridad"
navActive = "profile"
RequireLogin()

Dim rsUser, action, errMsg, okMsg
Dim fEmail, fBirthday, fPhoto, fUser
Dim curPass, newPass, newPass2, photoCsv, photoArr, i, plCount, songCount, birthVal

action = FM("action")
errMsg = ""

Set rsUser = DbQuery("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(CurrentUserId()))
If rsUser.EOF Then
    rsUser.Close
    Call Flash("No se pudo cargar tu perfil.", "error")
    GoToPage AppBase() & "home.asp"
End If

fUser = rsUser("USERNAME") & ""
fEmail = rsUser("EMAIL") & ""
fBirthday = FormatDateInput(rsUser("BIRTHDAY"))
fPhoto = rsUser("USER_PHOTO") & ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not CsrfValid() Then
        errMsg = "La sesión expiró. Vuelve a intentarlo."
    ElseIf action = "profile" Then
        fUser = FM("username")
        fEmail = FM("email")
        fBirthday = FM("birthday")
        fPhoto = FM("photo")

        If fUser = "" Or Len(fUser) < 3 Then
            errMsg = "El nombre de usuario debe tener al menos 3 caracteres."
        ElseIf Len(fUser) > 50 Then
            errMsg = "El nombre de usuario no puede superar los 50 caracteres."
        ElseIf ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", Array(fUser, CurrentUserId())), 0) > 0 Then
            errMsg = "Ese nombre de usuario ya está en uso."
        ElseIf Not IsValidEmail(fEmail) Then
            errMsg = "Ingresa un correo electronico válido."
        ElseIf ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", Array(fEmail, CurrentUserId())), 0) > 0 Then
            errMsg = "Ese correo ya está registrado por otro usuario."
        ElseIf fBirthday <> "" And IsEmpty(IsoDate(fBirthday)) Then
            errMsg = "La fecha de nacimiento no es válida."
        Else
            If fBirthday = "" Then
                birthVal = Null
            Else
                birthVal = IsoDate(fBirthday)
            End If
            Call DbExecute("UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ? WHERE ID_USER = ?", _
                Array(fUser, fEmail, birthVal, fPhoto, CurrentUserId()))
            Call RefreshUserSession()
            Call Flash("Tus datos se actualizaron correctamente.", "success")
            GoToPage AppBase() & "profile.asp"
        End If

    ElseIf action = "password" Then
        curPass = FM("current")
        newPass = FM("new")
        newPass2 = FM("confirm")

        If curPass = "" Or newPass = "" Then
            errMsg = "Completa todos los campos de contraseña."
        ElseIf Len(newPass) < 6 Then
            errMsg = "La nueva contraseña debe tener al menos 6 caracteres."
        ElseIf Len(newPass) > 50 Then
            errMsg = "La nueva contraseña no puede superar los 50 caracteres."
        ElseIf newPass <> newPass2 Then
            errMsg = "La confirmación no coincide con la nueva contraseña."
        ElseIf ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE ID_USER = ? AND (PASSWORD = ? OR UPPER(PASSWORD) = UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2)))", _
                Array(CurrentUserId(), curPass, curPass)), 0) = 0 Then
            errMsg = "La contraseña actual no es correcta."
        Else
            Call DbExecute("UPDATE TBL_USER SET PASSWORD = CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2) WHERE ID_USER = ?", _
                Array(newPass, CurrentUserId()))
            Call Flash("Tu contraseña se actualizo correctamente.", "success")
            GoToPage AppBase() & "profile.asp"
        End If
    End If
End If

photoCsv = ListFilesCsv(DIR_USERS, Array("jpg", "jpeg", "png", "webp", "gif"))
If photoCsv = "" Then photoCsv = "~/userImages/default.jpg"
photoArr = Split(photoCsv, "|")

plCount = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(CurrentUserId())), 0)
songCount = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps INNER JOIN TBL_PLAYLIST p ON p.ID_PLAYLIST = ps.ID_PLAYLIST WHERE p.ID_USER = ?", Array(CurrentUserId())), 0)
%>
<!--#include file="includes/layout_top.asp"-->

<section class="detail" style="align-items:center">
    <div class="detail__media" style="width:150px;flex-basis:150px">
        <%=Avatar(rsUser("USER_PHOTO") & "", rsUser("USERNAME") & "", "avatar--xl")%>
    </div>
    <div class="detail__body">
        <div class="detail__kind">Cuenta de <%=RoleLabel(rsUser("USER_TYPE"))%></div>
        <h2><%=HtmlEncode(rsUser("USERNAME") & "")%></h2>
        <div class="detail__meta">
            <span><%=Icon("mail", "")%>&#32;<%=HtmlEncode(rsUser("EMAIL") & "")%></span>
            <span class="dot"></span>
            <span><%=Icon("calendar", "")%>&#32;<%=IIf(IsNull(rsUser("BIRTHDAY")), "Sin fecha de nacimiento", FormatDateLong(rsUser("BIRTHDAY")))%></span>
        </div>
        <div class="detail__meta" style="margin-top:12px">
            <span class="chip"><%=Icon("list", "")%>&#32;<%=plCount%> playlists</span>
            <span class="chip"><%=Icon("music", "")%>&#32;<%=songCount%> canciones guardadas</span>
        </div>
    </div>
</section>

<% If errMsg <> "" Then %>
<div class="alert alert--error mt-24"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
<% End If %>

<div class="grid grid--two mt-24">
    <section class="panel">
        <div class="panel__head">
            <h3><%=Icon("user", "")%> Datos personales</h3>
        </div>
        <form method="post" action="<%=AppBase()%>profile.asp">
            <%=CsrfField()%>
            <input type="hidden" name="action" value="profile">
            <div class="form-grid">
                <div class="field">
                    <label for="pfUser">Nombre de usuario</label>
                    <input class="input" id="pfUser" name="username" maxlength="50" required value="<%=HtmlAttr(fUser)%>">
                </div>
                <div class="field">
                    <label for="pfEmail">Correo electronico</label>
                    <input class="input" id="pfEmail" name="email" type="email" maxlength="50" required value="<%=HtmlAttr(fEmail)%>">
                </div>
                <div class="field">
                    <label for="pfBirth">Fecha de nacimiento</label>
                    <input class="input" id="pfBirth" name="birthday" type="date" value="<%=HtmlAttr(fBirthday)%>">
                </div>
                <div class="field">
                    <label>Rol</label>
                    <input class="input" value="<%=RoleLabel(rsUser("USER_TYPE"))%>" readonly>
                    <span class="field__hint">El rol solo puede cambiarlo un administrador.</span>
                </div>
                <div class="field field--full">
                    <label>Foto de perfil</label>
                    <div class="pick-grid">
                        <% For i = 0 To UBound(photoArr) %>
                        <input type="radio" id="uf<%=i%>" name="photo" value="<%=HtmlAttr(photoArr(i))%>"<%=IIf(photoArr(i) = fPhoto, " checked", "")%>>
                        <label for="uf<%=i%>"><%=Cover(photoArr(i), "Foto " & (i + 1), "cover--rounded")%></label>
                        <% Next %>
                    </div>
                    <span class="field__hint">Las imágenes disponibles se leen de la carpeta <b>userImages</b>.</span>
                </div>
            </div>
            <div class="form-actions">
                <button class="btn btn--primary" type="submit"><%=Icon("save", "")%> Guardar cambios</button>
            </div>
        </form>
    </section>

    <section class="panel">
        <div class="panel__head">
            <h3><%=Icon("lock", "")%> Seguridad</h3>
        </div>
        <form method="post" action="<%=AppBase()%>profile.asp">
            <%=CsrfField()%>
            <input type="hidden" name="action" value="password">
            <div class="stack">
                <div class="field">
                    <label for="pwCur">Contraseña actual</label>
                    <input class="input" id="pwCur" name="current" type="password" required autocomplete="current-password">
                </div>
                <div class="field">
                    <label for="pwNew">Nueva contraseña</label>
                    <input class="input" id="pwNew" name="new" type="password" required minlength="6" maxlength="50" autocomplete="new-password">
                    <span class="field__hint">Mínimo 6 caracteres. Se almacena como hash SHA-1.</span>
                </div>
                <div class="field">
                    <label for="pwNew2">Confirmar nueva contraseña</label>
                    <input class="input" id="pwNew2" name="confirm" type="password" required minlength="6" maxlength="50" autocomplete="new-password">
                </div>
            </div>
            <div class="form-actions">
                <button class="btn btn--soft" type="submit"><%=Icon("lock", "")%> Actualizar contraseña</button>
            </div>
        </form>
    </section>
</div>

<% rsUser.Close %>
<!--#include file="includes/layout_bottom.asp"-->
