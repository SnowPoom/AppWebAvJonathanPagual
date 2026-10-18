<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="../includes/config.asp"-->
<!--#include file="../includes/functions.asp"-->
<!--#include file="../includes/auth.asp"-->
<!--#include file="../includes/ui.asp"-->
<%
Dim pageTitle, pageSubtitle, navActive
pageTitle = "Usuarios"
pageSubtitle = "Gestion de cuentas y roles"
navActive = "admin_users"
RequireAdmin()

Dim q, rs, action, idUser, targetType, targetName, conn, errMsg

action = FM("action")
errMsg = ""

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    If Not CsrfValid() Then
        Call Flash("La sesión expiró. Vuelve a intentarlo.", "error")
        GoToPage AppBase() & "admin/users.asp"
    ElseIf action = "delete" Then
        idUser = ToInt(FM("id"), 0)
        targetType = ToInt(DbScalar("SELECT USER_TYPE FROM TBL_USER WHERE ID_USER = ?", Array(idUser)), -1)
        targetName = DbScalar("SELECT USERNAME FROM TBL_USER WHERE ID_USER = ?", Array(idUser)) & ""

        If idUser = CurrentUserId() Then
            Call Flash("No puedes eliminar tu propia cuenta mientras la estas usando.", "error")
        ElseIf targetType < 0 Then
            Call Flash("El usuario indicado no existe.", "error")
        ElseIf targetType = ROLE_ADMIN And ToInt(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = 1", Empty), 0) <= 1 Then
            Call Flash("Debe existir al menos un administrador en el sistema.", "error")
        Else
            Set conn = Db()
            On Error Resume Next
            conn.BeginTrans
            Call DbExecute("DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = ?)", Array(idUser))
            Call DbExecute("DELETE FROM TBL_PLAYLIST WHERE ID_USER = ?", Array(idUser))
            Call DbExecute("DELETE FROM TBL_USER WHERE ID_USER = ?", Array(idUser))
            If Err.Number <> 0 Then
                conn.RollbackTrans
                errMsg = "No se pudo eliminar el usuario: " & Err.Description
                Err.Clear
            Else
                conn.CommitTrans
                Call Flash("Usuario '" & targetName & "' eliminado junto con sus playlists.", "success")
            End If
            On Error GoTo 0
        End If
        GoToPage AppBase() & "admin/users.asp"
    End If
End If

q = QS("q")
If q <> "" Then
    Set rs = DbQuery("SELECT u.ID_USER, u.USERNAME, u.EMAIL, u.BIRTHDAY, u.USER_PHOTO, u.USER_TYPE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS PLAYLISTS " & _
        "FROM TBL_USER u WHERE u.USERNAME LIKE ? OR u.EMAIL LIKE ? ORDER BY u.ID_USER", Array(LikeParam(q), LikeParam(q)))
Else
    Set rs = DbQuery("SELECT u.ID_USER, u.USERNAME, u.EMAIL, u.BIRTHDAY, u.USER_PHOTO, u.USER_TYPE, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS PLAYLISTS " & _
        "FROM TBL_USER u ORDER BY u.ID_USER", Empty)
End If
%>
<!--#include file="../includes/layout_top.asp"-->

<% If errMsg <> "" Then %>
<div class="alert alert--error"><%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span></div>
<% End If %>

<div class="admin-head">
    <div>
        <h2><%=rs.RecordCount%>&#32;<%=IIf(rs.RecordCount = 1, "cuenta registrada", "cuentas registradas")%></h2>
        <p>Roles definidos en TBL_CATALOG: 0 = Usuario, 1 = Administrador.</p>
    </div>
    <div class="row">
        <form method="get" action="<%=AppBase()%>admin/users.asp" class="row" style="gap:6px">
            <input class="input" type="search" name="q" value="<%=HtmlAttr(q)%>" placeholder="Buscar usuario o correo" style="height:40px;min-width:200px">
            <button class="btn btn--ghost btn--sm" type="submit"><%=Icon("search", "")%></button>
        </form>
        <a class="btn btn--primary" href="<%=AppBase()%>admin/user_form.asp"><%=Icon("plus", "")%> Nuevo usuario</a>
    </div>
</div>

<div class="table-wrap">
    <table class="data">
        <thead>
            <tr>
                <th>Usuario</th>
                <th>Correo</th>
                <th>Nacimiento</th>
                <th>Playlists</th>
                <th>Rol</th>
                <th style="text-align:right">Acciones</th>
            </tr>
        </thead>
        <tbody>
            <% Do Until rs.EOF %>
            <tr>
                <td>
                    <div class="cell-user">
                        <%=Avatar(rs("USER_PHOTO") & "", rs("USERNAME") & "", "avatar--sm")%>
                        <div>
                            <strong><%=HtmlEncode(rs("USERNAME") & "")%></strong>
                            <div class="small muted">ID <%=rs("ID_USER")%><% If ToInt(rs("ID_USER"), 0) = CurrentUserId() Then %> &middot; tu cuenta<% End If %></div>
                        </div>
                    </div>
                </td>
                <td class="small"><%=HtmlEncode(rs("EMAIL") & "")%></td>
                <td class="small"><%=FormatDateShort(rs("BIRTHDAY"))%></td>
                <td class="num"><%=ToInt(rs("PLAYLISTS"), 0)%></td>
                <td><span class="chip<%=IIf(ToInt(rs("USER_TYPE"), 0) = 1, " chip--admin", "")%>"><%=RoleLabel(rs("USER_TYPE"))%></span></td>
                <td>
                    <div class="cell-actions">
                        <a class="icon-btn" href="<%=AppBase()%>admin/user_form.asp?id=<%=rs("ID_USER")%>" title="Editar"><%=Icon("edit", "")%></a>
                        <% If ToInt(rs("ID_USER"), 0) <> CurrentUserId() Then %>
                        <form method="post" action="<%=AppBase()%>admin/users.asp" style="display:inline">
                            <%=CsrfField()%>
                            <input type="hidden" name="action" value="delete">
                            <input type="hidden" name="id" value="<%=rs("ID_USER")%>">
                            <button class="icon-btn icon-btn--danger" type="submit" title="Eliminar"
                                    data-confirm="Se eliminara el usuario y todas sus playlists. Esta acción no se puede deshacer."
                                    data-confirm-title="Eliminar usuario" data-confirm-ok="Si, eliminar"><%=Icon("trash", "")%></button>
                        </form>
                        <% End If %>
                    </div>
                </td>
            </tr>
            <% rs.MoveNext : Loop %>
        </tbody>
    </table>
</div>

<% rs.Close %>
<!--#include file="../includes/layout_bottom.asp"-->
