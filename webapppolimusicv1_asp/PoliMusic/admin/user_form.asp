<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
RequireAdmin()

Dim idUser, isEdit, rs, errMsg, action
Dim fUser, fEmail, fPass, fBirth, fPhoto, fType, photoCsv, photoArr, i, birthVal, adminCount, plCount

idUser = ToInt(QS("id"), 0)
isEdit = (idUser > 0)
errMsg = ""

fUser = ""
fEmail = ""
fPass = ""
fBirth = ""
fPhoto = "~/userImages/default.jpg"
fType = ROLE_USER
adminCount = 0
plCount = 0

If isEdit Then
    Set rs = DbQuery("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(idUser))
    If rs.EOF Then
        rs.Close
        Call Flash("El usuario solicitado no existe.", "error")
        GoToPage AppBase() & "admin/users.asp"
    End If
    fUser = rs("USERNAME") & ""
    fEmail = rs("EMAIL") & ""
    fBirth = FormatDateInput(rs("BIRTHDAY"))
    fPhoto = rs("USER_PHOTO") & ""
    fType = ToInt(rs("USER_TYPE"), ROLE_USER)
    adminCount = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1", Empty), 0)
    plCount = ToInt(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(idUser)), 0)
End If

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    fUser = FM("username")
    fEmail = FM("email")
    fPass = FM("password")
    fBirth = FM("birthday")
    fPhoto = FM("photo")
    fType = ToInt(FM("type"), ROLE_USER)
    If fType <> ROLE_ADMIN Then fType = ROLE_USER
    If fPhoto = "" Then fPhoto = "~/userImages/default.jpg"

    If Not CsrfValid() Then
        errMsg = "La sesión expiró. Vuelve a intentarlo."
    ElseIf fUser = "" Or Len(fUser) < 3 Then
        errMsg = "El nombre de usuario debe tener al menos 3 caracteres."
    ElseIf Len(fUser) > 50 Then
        errMsg = "El nombre de usuario no puede superar los 50 caracteres."
    ElseIf Not IsValidEmail(fEmail) Then
        errMsg = "Ingresa un correo electronico válido."
    ElseIf fBirth <> "" And IsEmpty(IsoDate(fBirth)) Then
        errMsg = "La fecha de nacimiento no es válida."
    ElseIf Not isEdit And Len(fPass) < 6 Then
        errMsg = "La contraseña debe tener al menos 6 caracteres."
    ElseIf isEdit And fPass <> "" And Len(fPass) < 6 Then
        errMsg = "La nueva contraseña debe tener al menos 6 caracteres."
    ElseIf isEdit And fType = ROLE_USER And ToInt(DbScalar("SELECT USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(idUser)), 0) = ROLE_ADMIN And adminCount <= 1 Then
        errMsg = "No puedes quitar el rol de administrador al último administrador."
    Else
        If fBirth = "" Then
            birthVal = Null
        Else
            birthVal = IsoDate(fBirth)
        End If

        If isEdit Then
            If ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ? AND ID_USER <> ?", Array(fUser, idUser)), 0) > 0 Then
                errMsg = "Ese nombre de usuario ya está en uso."
            ElseIf ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ? AND ID_USER <> ?", Array(fEmail, idUser)), 0) > 0 Then
                errMsg = "Ese correo ya está registrado."
            Else
                If fPass = "" Then
                    Call DbExecute("UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?", _
                        Array(fUser, fEmail, birthVal, fPhoto, fType, idUser))
                Else
                    Call DbExecute("UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ?, " & _
                        "PASSWORD = CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2) WHERE ID_USER = ?", _
                        Array(fUser, fEmail, birthVal, fPhoto, fType, fPass, idUser))
                End If
                If idUser = CurrentUserId() Then Call RefreshUserSession()
                Call Flash("Usuario '" & fUser & "' actualizado correctamente.", "success")
                GoToPage AppBase() & "admin/users.asp"
            End If
        Else
            If ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = ?", Array(fUser)), 0) > 0 Then
                errMsg = "Ese nombre de usuario ya está en uso."
            ElseIf ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = ?", Array(fEmail)), 0) > 0 Then
                errMsg = "Ese correo ya está registrado."
            Else
                Call DbExecute("INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
                    "VALUES (?, CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2), ?, ?, ?, ?)", _
                    Array(fUser, fPass, fEmail, birthVal, fPhoto, fType))
                Call Flash("Usuario '" & fUser & "' creado correctamente.", "success")
                GoToPage AppBase() & "admin/users.asp"
            End If
        End If
    End If
End If

pageTitle = IIf(isEdit, "Editar usuario", "Nuevo usuario")
pageSubtitle = IIf(isEdit, "Actualiza los datos de la cuenta", "Registra una cuenta en PoliMusic")
navActive = "admin_users"

photoCsv = ListFilesCsv(DIR_USERS, Array("jpg", "jpeg", "png", "webp", "gif"))
If photoCsv = "" Then photoCsv = "~/userImages/default.jpg"
photoArr = Split(photoCsv, "|")
%>
<!--#include file="../includes/layout_top.asp"-->

<a class="btn btn--ghost btn--sm" href="<%=AppBase()%>admin/users.asp" style="margin-bottom:16px"><%=Icon("arrow-left", "")%> Volver a usuarios</a>

<section class="panel">
    <div class="panel__head">
        <h3><%=Icon("user", "")%>&#32;<%=HtmlEncode(pageTitle)%></h3>
        <% If isEdit Then %>
        <span class="chip chip--muted">ID <%=idUser%> &middot; <%=plCount%> playlists</span>
        <% End If %>
    </div>

    <% If errMsg <> "" Then %>
    <div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
    <% End If %>

    <form method="post" action="<%=AppBase()%>admin/user_form.asp<%=IIf(isEdit, "?id=" & idUser, "")%>">
        <%=CsrfField()%>
        <div class="form-grid">
            <div class="field">
                <label for="uName">Nombre de usuario</label>
                <input class="input" id="uName" name="username" maxlength="50" required value="<%=HtmlAttr(fUser)%>">
            </div>
            <div class="field">
                <label for="uEmail">Correo electronico</label>
                <input class="input" id="uEmail" name="email" type="email" maxlength="50" required value="<%=HtmlAttr(fEmail)%>">
            </div>
            <div class="field">
                <label for="uPass"><%=IIf(isEdit, "Nueva contraseña (opcional)", "Contraseña")%></label>
                <input class="input" id="uPass" name="password" type="password" maxlength="50" <%=IIf(isEdit, "", "required")%>
                       autocomplete="new-password" placeholder="<%=IIf(isEdit, "Dejar en blanco para conservar", "Mínimo 6 caracteres")%>">
                <span class="field__hint">Se guarda como hash SHA-1 calculado por SQL Server.</span>
            </div>
            <div class="field">
                <label for="uType">Rol</label>
                <select class="select" id="uType" name="type">
                    <option value="0"<%=IIf(fType = ROLE_USER, " selected", "")%>>Usuario</option>
                    <option value="1"<%=IIf(fType = ROLE_ADMIN, " selected", "")%>>Administrador</option>
                </select>
            </div>
            <div class="field">
                <label for="uBirth">Fecha de nacimiento</label>
                <input class="input" id="uBirth" name="birthday" type="date" value="<%=HtmlAttr(fBirth)%>">
            </div>
            <div class="field field--full">
                <label>Foto de perfil</label>
                <div class="pick-grid">
                    <% For i = 0 To UBound(photoArr) %>
                    <input type="radio" id="uf<%=i%>" name="photo" value="<%=HtmlAttr(photoArr(i))%>"<%=IIf(photoArr(i) = fPhoto, " checked", "")%>>
                    <label for="uf<%=i%>"><%=Cover(photoArr(i), "Foto " & (i + 1), "cover--rounded")%></label>
                    <% Next %>
                </div>
                <span class="field__hint">Imágenes disponibles en la carpeta <b>userImages</b>.</span>
            </div>
        </div>
        <div class="form-actions">
            <button class="btn btn--primary" type="submit"><%=Icon("save", "")%>&#32;<%=IIf(isEdit, "Guardar cambios", "Crear usuario")%></button>
            <a class="btn btn--ghost" href="<%=AppBase()%>admin/users.asp">Cancelar</a>
        </div>
    </form>
</section>

<% If isEdit Then rs.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
