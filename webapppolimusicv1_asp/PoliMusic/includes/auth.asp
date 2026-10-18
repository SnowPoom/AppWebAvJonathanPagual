<%
' =====================================================================
'  PoliMusic - Autenticacion, sesión y control de roles
'  Roles: TBL_USER.USER_TYPE -> 0 = Usuario, 1 = Administrador
'         (valores definidos en TBL_CATALOG, grupo 1)
' =====================================================================

Function IsLoggedIn()
    IsLoggedIn = (Trim(Session("UsuarioID") & "") <> "")
End Function

Function CurrentUserId()
    CurrentUserId = ToInt(Session("UsuarioID"), 0)
End Function

Function CurrentUserName()
    CurrentUserName = Session("Usuario") & ""
End Function

Function IsAdmin()
    IsAdmin = (ToInt(Session("UsuarioTipo"), -1) = ROLE_ADMIN)
End Function

Function RoleLabel(ByVal userType)
    If ToInt(userType, 0) = ROLE_ADMIN Then
        RoleLabel = "Administrador"
    Else
        RoleLabel = "Usuario"
    End If
End Function

' Carga los datos del usuario en la sesión (se llama desde login.asp)
Sub StartUserSession(ByVal idUser, ByVal userName, ByVal userType, ByVal userPhoto, ByVal userEmail)
    Session("UsuarioID")    = idUser
    Session("Usuario")      = userName
    Session("UsuarioTipo")  = userType
    Session("UsuarioFoto")  = userPhoto
    Session("UsuarioEmail") = userEmail
    Session("login_time")   = Now()
End Sub

Sub EndUserSession()
    Session.Contents.RemoveAll()
    Session.Abandon()
End Sub

' ---------------------------------------------------------------------
'  Guardas de acceso
' ---------------------------------------------------------------------
Sub RequireLogin()
    Dim target
    If Not IsLoggedIn() Then
        target = AppBase() & "login.asp?next=" & Server.URLEncode(CurrentPath())
        GoToPage target
    End If
End Sub

Sub RequireAdmin()
    RequireLogin()
    If Not IsAdmin() Then
        Flash "No tienes permisos para acceder a esa sección.", "error"
        GoToPage AppBase() & "home.asp"
    End If
End Sub

' Refresca en sesión los datos del usuario (tras editar el perfil)
Sub RefreshUserSession()
    Dim rs
    If Not IsLoggedIn() Then Exit Sub
    Set rs = DbQuery("SELECT USERNAME, USER_TYPE, USER_PHOTO, EMAIL FROM TBL_USER WHERE ID_USER = ?", Array(CurrentUserId()))
    If Not rs.EOF Then
        Session("Usuario")      = rs("USERNAME") & ""
        Session("UsuarioTipo")  = ToInt(rs("USER_TYPE"), ROLE_USER)
        Session("UsuarioFoto")  = rs("USER_PHOTO") & ""
        Session("UsuarioEmail") = rs("EMAIL") & ""
    End If
    rs.Close
    Set rs = Nothing
End Sub

' ---------------------------------------------------------------------
'  Verificacion de credenciales
'  La BD contiene contraseñas en texto plano (semilla) y en SHA-1 (40 hex).
'  Se acepta cualquiera de las dos formas y las nuevas se guardan con SHA-1
'  calculado por SQL Server (HASHBYTES), sin depender de componentes COM.
' ---------------------------------------------------------------------
Function FindUserForLogin(ByVal userName, ByVal plainPassword)
    Dim sql, rs
    sql = "SELECT TOP 1 ID_USER, USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE " & _
          "FROM TBL_USER WHERE USERNAME = ? AND (PASSWORD = ? OR UPPER(PASSWORD) = " & _
          "UPPER(CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2)))"
    Set rs = DbQuery(sql, Array(userName, plainPassword, plainPassword))
    Set FindUserForLogin = rs
End Function

' Hash SHA-1 en hexadecimal mayusculas, calculado por el motor SQL
Function Sha1Hex(ByVal plainPassword)
    Dim v
    v = DbScalar("SELECT CONVERT(varchar(40), HASHBYTES('SHA1', CAST(? AS varchar(50))), 2)", Array(plainPassword))
    Sha1Hex = UCase(v & "")
End Function
%>
