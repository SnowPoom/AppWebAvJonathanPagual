<%
' =====================================================================
'  PoliMusic - Cabecera y estructura comun (layout)
'  Variables esperadas: pageTitle, pageSubtitle, navActive
' =====================================================================
Dim flData, flKind, flMsg, sessUser, sessPhoto, roleTxt, userInitials
flData = TakeFlash()
flKind = "info"
flMsg = ""
If flData <> "" Then
    If InStr(flData, "|") > 0 Then
        flKind = Left(flData, InStr(flData, "|") - 1)
        flMsg = Mid(flData, InStr(flData, "|") + 1)
    Else
        flMsg = flData
    End If
End If
sessUser = Session("Usuario") & ""
sessPhoto = Session("UsuarioFoto") & ""
roleTxt = RoleLabel(Session("UsuarioTipo"))
userInitials = Initials(sessUser)
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><%=HtmlEncode(pageTitle)%> &middot; <%=APP_NAME%></title>
<meta name="description" content="PoliMusic: plataforma de música con identidad visual propia.">
<meta name="theme-color" content="#04121f">
<link rel="icon" type="image/svg+xml" href="<%=AppBase()%>assets/img/favicon.svg">
<link rel="stylesheet" href="<%=AppBase()%>assets/css/fonts.css?v=<%=APP_VERSION%>">
<link rel="stylesheet" href="<%=AppBase()%>assets/css/polimusic.css?v=<%=APP_VERSION%>">
</head>
<body data-page="<%=HtmlAttr(navActive)%>">
<div class="bg" aria-hidden="true">
    <span class="bg__blob bg__blob--a"></span>
    <span class="bg__blob bg__blob--b"></span>
    <span class="bg__blob bg__blob--c"></span>
    <span class="bg__grid"></span>
    <span class="bg__noise"></span>
</div>

<div class="app">
    <aside class="sidebar glass" id="sidebar">
        <a class="brand" href="<%=AppBase()%>home.asp" aria-label="PoliMusic, ir al inicio">
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

        <nav class="nav" aria-label="Navegacion principal">
            <p class="nav__label">Descubrir</p>
            <a class="nav__item<%=IIf(navActive = "home", " is-active", "")%>" href="<%=AppBase()%>home.asp">
                <%=Icon("home", "")%><span>Inicio</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "songs", " is-active", "")%>" href="<%=AppBase()%>songs.asp">
                <%=Icon("music", "")%><span>Canciones</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "artists", " is-active", "")%>" href="<%=AppBase()%>artists.asp">
                <%=Icon("mic", "")%><span>Artistas</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "albums", " is-active", "")%>" href="<%=AppBase()%>albums.asp">
                <%=Icon("disc", "")%><span>Álbumes</span>
            </a>

            <p class="nav__label">Mi música</p>
            <a class="nav__item<%=IIf(navActive = "playlists", " is-active", "")%>" href="<%=AppBase()%>playlists.asp">
                <%=Icon("list", "")%><span>Mis playlists</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "profile", " is-active", "")%>" href="<%=AppBase()%>profile.asp">
                <%=Icon("user", "")%><span>Mi perfil</span>
            </a>

            <% If IsAdmin() Then %>
            <p class="nav__label">Administración</p>
            <a class="nav__item<%=IIf(navActive = "admin", " is-active", "")%>" href="<%=AppBase()%>admin/default.asp">
                <%=Icon("shield", "")%><span>Panel</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "admin_users", " is-active", "")%>" href="<%=AppBase()%>admin/users.asp">
                <%=Icon("users", "")%><span>Usuarios</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "admin_songs", " is-active", "")%>" href="<%=AppBase()%>admin/songs.asp">
                <%=Icon("music", "")%><span>Canciones</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "admin_artists", " is-active", "")%>" href="<%=AppBase()%>admin/artists.asp">
                <%=Icon("mic", "")%><span>Artistas</span>
            </a>
            <a class="nav__item<%=IIf(navActive = "admin_albums", " is-active", "")%>" href="<%=AppBase()%>admin/albums.asp">
                <%=Icon("disc", "")%><span>Álbumes</span>
            </a>
            <% End If %>
        </nav>

        <div class="sidebar__foot">
            <div class="user-card">
                <%=Avatar(sessPhoto, sessUser, "avatar--sm")%>
                <div class="user-card__meta">
                    <strong title="<%=HtmlAttr(sessUser)%>"><%=HtmlEncode(sessUser)%></strong>
                    <span class="chip chip--role<%=IIf(IsAdmin(), " chip--admin", "")%>"><%=roleTxt%></span>
                </div>
                <a class="icon-btn icon-btn--ghost" href="<%=AppBase()%>logout.asp" title="Cerrar sesión" aria-label="Cerrar sesión">
                    <%=Icon("logout", "")%>
                </a>
            </div>
        </div>
    </aside>

    <div class="scrim" id="scrim" hidden></div>

    <div class="app-main">
        <header class="topbar">
            <button class="icon-btn topbar__burger" id="navToggle" type="button" aria-label="Abrir menu">
                <%=Icon("menu", "")%>
            </button>
            <div class="topbar__titles">
                <h1><%=HtmlEncode(pageTitle)%></h1>
                <% If pageSubtitle & "" <> "" Then %><p><%=HtmlEncode(pageSubtitle)%></p><% End If %>
            </div>
            <form class="searchbar" action="<%=AppBase()%>search.asp" method="get" role="search">
                <%=Icon("search", "searchbar__ico")%>
                <input type="search" name="q" id="globalSearch" placeholder="Buscar canciones, artistas o álbumes"
                       value="<%=HtmlAttr(QS("q"))%>" autocomplete="off" aria-label="Buscar">
                <kbd>/</kbd>
            </form>
            <div class="usermenu">
                <button class="usermenu__trigger" id="userChip" type="button" aria-haspopup="true" aria-expanded="false">
                    <%=Avatar(sessPhoto, sessUser, "avatar--xs")%>
                    <span class="usermenu__name"><%=HtmlEncode(sessUser)%></span>
                    <%=Icon("chevron-down", "usermenu__caret")%>
                </button>
                <div class="dropdown" id="userDropdown" hidden>
                    <div class="dropdown__head">
                        <strong><%=HtmlEncode(sessUser)%></strong>
                        <span><%=HtmlEncode(Session("UsuarioEmail") & "")%></span>
                    </div>
                    <a class="dropdown__item" href="<%=AppBase()%>profile.asp"><%=Icon("user", "")%>Mi perfil</a>
                    <% If IsAdmin() Then %>
                    <a class="dropdown__item" href="<%=AppBase()%>admin/default.asp"><%=Icon("shield", "")%>Panel de administración</a>
                    <% End If %>
                    <div class="dropdown__sep"></div>
                    <a class="dropdown__item dropdown__item--danger" href="<%=AppBase()%>logout.asp"><%=Icon("logout", "")%>Cerrar sesión</a>
                </div>
            </div>
        </header>

        <main class="content" id="content">
            <% If flMsg <> "" Then %>
            <div class="alert alert--<%=HtmlAttr(flKind)%>" role="status">
                <%=Icon(IIf(flKind = "error", "alert", "check"), "")%>
                <span><%=HtmlEncode(flMsg)%></span>
                <button class="alert__close" type="button" aria-label="Cerrar">&times;</button>
            </div>
            <% End If %>
