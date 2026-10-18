<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<%
Dim errMsg, userName, userPass, nextUrl, attempts, lockUntil, rs

errMsg = ""
userName = ""
userPass = ""
nextUrl = SafeNext(QS("next"))
If nextUrl = "" Then nextUrl = SafeNext(FM("next"))

If IsLoggedIn() Then
    GoToPage AppBase() & "home.asp"
End If

attempts = ToInt(Session("login_attempts"), 0)
lockUntil = Session("login_lock_until")
If IsEmpty(lockUntil) Or Not IsDate(lockUntil) Then lockUntil = Empty

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    userName = FM("username")
    userPass = FM("password")

    If Not CsrfValid() Then
        errMsg = "La sesión expiró. Intenta nuevamente."
    ElseIf IsDate(lockUntil) And CDate(lockUntil) > Now() Then
        errMsg = "Demasiados intentos fallidos. Espera un momento antes de volver a intentar."
    ElseIf userName = "" Or userPass = "" Then
        errMsg = "Ingresa tu usuario y tu contraseña."
    Else
        Set rs = FindUserForLogin(userName, userPass)
        If rs.EOF Then
            attempts = attempts + 1
            Session("login_attempts") = attempts
            If attempts >= 5 Then
                Session("login_lock_until") = DateAdd("s", 30, Now())
                Session("login_attempts") = 0
                errMsg = "Demasiados intentos fallidos. La cuenta se bloqueo por 30 segundos."
            Else
                errMsg = "Usuario o contraseña incorrectos. Intento " & attempts & " de 5."
            End If
        Else
            Session("login_attempts") = 0
            Session("login_lock_until") = Empty
            StartUserSession rs("ID_USER"), rs("USERNAME") & "", ToInt(rs("USER_TYPE"), ROLE_USER), _
                             rs("USER_PHOTO") & "", rs("EMAIL") & ""
            rs.Close
            Set rs = Nothing
            If nextUrl <> "" Then
                GoToPage nextUrl
            Else
                GoToPage AppBase() & "home.asp"
            End If
        End If
        rs.Close
        Set rs = Nothing
    End If
End If
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Iniciar sesión &middot; <%=APP_NAME%></title>
<meta name="theme-color" content="#04121f">
<link rel="icon" type="image/svg+xml" href="<%=AppBase()%>assets/img/favicon.svg">
<link rel="stylesheet" href="<%=AppBase()%>assets/css/fonts.css?v=<%=APP_VERSION%>">
<link rel="stylesheet" href="<%=AppBase()%>assets/css/polimusic.css?v=<%=APP_VERSION%>">
</head>
<body>
<div class="bg" aria-hidden="true">
    <span class="bg__blob bg__blob--a"></span>
    <span class="bg__blob bg__blob--b"></span>
    <span class="bg__blob bg__blob--c"></span>
    <span class="bg__grid"></span>
    <span class="bg__noise"></span>
</div>

<div class="auth">
    <aside class="auth__aside">
        <div>
            <a class="brand" href="<%=AppBase()%>login.asp" style="padding:0">
                <span class="brand__mark">
                    <svg viewBox="0 0 40 40" aria-hidden="true">
                        <defs>
                            <linearGradient id="bm1" x1="0" y1="0" x2="1" y2="1">
                                <stop offset="0" stop-color="#5eead4"/><stop offset=".5" stop-color="#22d3ee"/><stop offset="1" stop-color="#3b82f6"/>
                            </linearGradient>
                        </defs>
                        <rect x="1.5" y="1.5" width="37" height="37" rx="12" fill="url(#bm1)" opacity=".16"/>
                        <rect x="1.5" y="1.5" width="37" height="37" rx="12" fill="none" stroke="url(#bm1)" stroke-width="1.4" opacity=".7"/>
                        <rect x="11" y="15" width="3.4" height="10" rx="1.7" fill="url(#bm1)"/>
                        <rect x="17" y="9.5" width="3.4" height="21" rx="1.7" fill="url(#bm1)"/>
                        <rect x="23" y="13" width="3.4" height="14" rx="1.7" fill="url(#bm1)"/>
                        <rect x="29" y="17.5" width="3.4" height="5" rx="1.7" fill="url(#bm1)"/>
                    </svg>
                </span>
                <span class="brand__text"><b>Poli</b><em>Music</em></span>
            </a>

            <h1 class="auth__title">Tu música,<br><span>en tu propio ritmo.</span></h1>
            <p class="auth__lead">Explora canciones, descubre artistas y construye playlists a tu medida en una plataforma diseñada para escuchar sin distracciones.</p>

            <div class="auth__features">
                <div class="auth__feature"><%=Icon("headphones", "")%><span>Reproductor completo con cola, volumen y progreso</span></div>
                <div class="auth__feature"><%=Icon("sparkles", "")%><span>Playlists personales públicas o privadas</span></div>
                <div class="auth__feature"><%=Icon("shield", "")%><span>Roles de administrador y usuario</span></div>
            </div>
        </div>
        <p class="auth__foot">&copy; <%=Year(Date())%>&#32;<%=APP_NAME%> &middot; Escuela Politécnica Nacional</p>
    </aside>

    <section class="auth__panel">
        <div class="auth__card">
            <h1>Bienvenido de nuevo</h1>
            <p>Ingresa con tu cuenta para continuar escuchando.</p>

            <% If errMsg <> "" Then %>
            <div class="alert alert--error" style="margin:18px 0 0">
                <%=Icon("alert", "")%><span><%=HtmlEncode(errMsg)%></span>
            </div>
            <% End If %>

            <form class="auth__form" method="post" action="<%=AppBase()%>login.asp" autocomplete="on">
                <%=CsrfField()%>
                <input type="hidden" name="next" value="<%=HtmlAttr(nextUrl)%>">
                <div class="field">
                    <label for="username">Usuario</label>
                    <input class="input" type="text" id="username" name="username" value="<%=HtmlAttr(userName)%>"
                           placeholder="Tu nombre de usuario" required maxlength="50" autofocus>
                </div>
                <div class="field">
                    <label for="password">Contraseña</label>
                    <div class="pw-wrap">
                        <input class="input" type="password" id="password" name="password"
                               placeholder="Tu contraseña" required maxlength="50">
                        <button class="pw-toggle" type="button" id="pwToggle" aria-label="Mostrar contraseña">
                            <span id="pwIcon"><%=Icon("eye", "")%></span>
                        </button>
                    </div>
                </div>
                <button class="btn btn--primary btn--block" type="submit" style="margin-top:4px">
                    <%=Icon("play", "")%> Iniciar sesión
                </button>
            </form>

            <div class="auth__demo">
                <b>Cuentas de prueba</b><br>
                Administrador: <code>userAdmin</code> / <code>12345678</code><br>
                Usuario: <code>userNormal</code> / <code>12345678</code>
            </div>
        </div>
    </section>
</div>

<script>
(function () {
    var t = document.getElementById('pwToggle');
    var p = document.getElementById('password');
    var i = document.getElementById('pwIcon');
    if (!t || !p) return;
    var eyeOpen = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2.6 12S6 6.4 12 6.4 21.4 12 21.4 12 18 17.6 12 17.6 2.6 12 2.6 12z"/><circle cx="12" cy="12" r="2.9"/></svg>';
    var eyeOff = '<svg class="ico" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 4.4 20 19.6"/><path d="M9.6 7.1A9.7 9.7 0 0 1 12 6.8c6 0 9.4 5.2 9.4 5.2a16 16 0 0 1-2.8 3.3"/><path d="M6.5 8.4A15.9 15.9 0 0 0 2.6 12s3.4 5.2 9.4 5.2a9.6 9.6 0 0 0 3.6-.7"/><path d="M10.2 10.4a2.6 2.6 0 0 0 3.5 3.6"/></svg>';
    t.addEventListener('click', function () {
        var show = p.type === 'password';
        p.type = show ? 'text' : 'password';
        t.setAttribute('aria-label', show ? 'Ocultar contraseña' : 'Mostrar contraseña');
        t.style.color = show ? '#5eead4' : '';
        if (i) i.innerHTML = show ? eyeOff : eyeOpen;
    });
})();
</script>
</body>
</html>
<% DbClose() %>
