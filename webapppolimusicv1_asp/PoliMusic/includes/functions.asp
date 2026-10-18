<%
' =====================================================================
'  PoliMusic - Funciones utilitarias compartidas
' =====================================================================

' ---------------------------------------------------------------------
'  Rutas de la aplicación
'  AppBase devuelve la raíz virtual del sitio tanto si PoliMusic se publica
'  en la raíz del sitio (/), como en una carpeta (/PoliMusic/) o como
'  aplicación IIS. Se deduce del script en ejecución y de las dos únicas
'  carpetas de sección existentes: admin y api.
' ---------------------------------------------------------------------
Function AppBase()
    Dim s, dirPath
    If gAppBase <> "" Then
        AppBase = gAppBase
        Exit Function
    End If

    s = Request.ServerVariables("SCRIPT_NAME") & ""
    If s = "" Then s = "/default.asp"
    dirPath = Left(s, InStrRev(s, "/"))

    If LCase(Right(dirPath, 7)) = "/admin/" Then
        dirPath = Left(dirPath, Len(dirPath) - 6)
    ElseIf LCase(Right(dirPath, 5)) = "/api/" Then
        dirPath = Left(dirPath, Len(dirPath) - 4)
    End If

    If dirPath = "" Then dirPath = "/"
    gAppBase = dirPath
    AppBase = gAppBase
End Function

' Convierte "~/songFiles/x.mp3" en "/PoliMusic/songFiles/x.mp3"
Function AssetUrl(ByVal p)
    Dim s
    s = Trim(p & "")
    If s = "" Then
        AssetUrl = ""
        Exit Function
    End If
    If LCase(Left(s, 5)) = "http:" Or LCase(Left(s, 6)) = "https:" Then
        AssetUrl = s
        Exit Function
    End If
    s = Replace(s, "\", "/")
    If Left(s, 2) = "~/" Then s = Mid(s, 3)
    Do While Left(s, 1) = "/"
        s = Mid(s, 2)
    Loop
    AssetUrl = AppBase() & s
End Function

Function Fso()
    If gFso Is Nothing Then Set gFso = Server.CreateObject("Scripting.FileSystemObject")
    Set Fso = gFso
End Function

' True si el archivo referenciado por la BD existe realmente en disco
Function AssetExists(ByVal p)
    Dim s, oFs
    AssetExists = False
    s = Trim(p & "")
    If s = "" Then Exit Function
    If LCase(Left(s, 5)) = "http:" Or LCase(Left(s, 6)) = "https:" Then
        AssetExists = True
        Exit Function
    End If
    On Error Resume Next
    Set oFs = Fso()
    If Not (oFs Is Nothing) Then AssetExists = oFs.FileExists(Server.MapPath(AssetUrl(s)))
    If Err.Number <> 0 Then AssetExists = False
    Err.Clear
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------
'  Salida segura
' ---------------------------------------------------------------------
Function HtmlEncode(ByVal s)
    Dim t
    t = CStr(s & "")
    t = Replace(t, "&", "&amp;")
    t = Replace(t, "<", "&lt;")
    t = Replace(t, ">", "&gt;")
    t = Replace(t, """", "&quot;")
    t = Replace(t, "'", "&#39;")
    HtmlEncode = t
End Function

Function HtmlAttr(ByVal s)
    HtmlAttr = HtmlEncode(s)
End Function

Function TextOr(ByVal s, ByVal fallback)
    If Trim(s & "") = "" Then
        TextOr = fallback
    Else
        TextOr = s
    End If
End Function

' Equivalente a IIf de VBA (VBScript no lo incluye)
Function IIf(ByVal cond, ByVal whenTrue, ByVal whenFalse)
    If cond Then
        IIf = whenTrue
    Else
        IIf = whenFalse
    End If
End Function

Function Truncate(ByVal s, ByVal maxLen)
    Dim t
    t = CStr(s & "")
    If Len(t) > maxLen Then
        Truncate = Left(t, maxLen - 1) & "..."
    Else
        Truncate = t
    End If
End Function

' Cadena JSON segura (incluye comillas)
Function JsonStr(ByVal s)
    Dim i, ch, code, out
    s = CStr(s & "")
    out = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case ch
            Case """"
                out = out & "\"""
            Case "\"
                out = out & "\\"
            Case "/"
                out = out & "\/"
            Case vbCr
                out = out & "\r"
            Case vbLf
                out = out & "\n"
            Case vbTab
                out = out & "\t"
            Case Else
                If code < 32 Then
                    out = out & "\u" & Right("000" & Hex(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next
    JsonStr = """" & out & """"
End Function

' ---------------------------------------------------------------------
'  Conversion de datos
' ---------------------------------------------------------------------
Function ToInt(ByVal v, ByVal def)
    Dim n
    On Error Resume Next
    n = CLng(v)
    If Err.Number <> 0 Then
        n = def
        Err.Clear
    End If
    On Error GoTo 0
    ToInt = n
End Function

Function QS(ByVal name)
    QS = Trim(Request.QueryString(name) & "")
End Function

Function FM(ByVal name)
    FM = Trim(Request.Form(name) & "")
End Function

Function IsValidEmail(ByVal mail)
    Dim re
    IsValidEmail = False
    If Len(mail) < 5 Or InStr(mail, "@") = 0 Then Exit Function
    Set re = New RegExp
    re.Pattern = "^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"
    re.IgnoreCase = True
    IsValidEmail = re.Test(mail)
    Set re = Nothing
End Function

' Patron LIKE seguro (escapa comodines de T-SQL)
Function LikeParam(ByVal s)
    s = CStr(s & "")
    s = Replace(s, "[", "[[]")
    s = Replace(s, "%", "[%]")
    s = Replace(s, "_", "[_]")
    LikeParam = "%" & s & "%"
End Function

' Acepta "YYYY-MM-DD" (input date) y devuelve "YYYY-MM-DDT00:00:00" o Empty
Function IsoDate(ByVal v)
    Dim y, m, d, parts
    IsoDate = Empty
    v = Trim(v & "")
    If Len(v) < 10 Then Exit Function
    parts = Split(Left(v, 10), "-")
    If UBound(parts) <> 2 Then Exit Function
    y = ToInt(parts(0), 0) : m = ToInt(parts(1), 0) : d = ToInt(parts(2), 0)
    If y < 1900 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    IsoDate = Right("0000" & y, 4) & "-" & Right("0" & m, 2) & "-" & Right("0" & d, 2) & "T00:00:00"
End Function

Function MonthNameEs(ByVal m)
    Dim names
    names = Array("", "enero", "febrero", "marzo", "abril", "mayo", "junio", _
                  "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")
    If m >= 1 And m <= 12 Then MonthNameEs = names(m) Else MonthNameEs = ""
End Function

Function FormatDateLong(ByVal d)
    If IsNull(d) Or Not IsDate(d) Then
        FormatDateLong = "Sin fecha"
        Exit Function
    End If
    FormatDateLong = Day(d) & " de " & MonthNameEs(Month(d)) & " de " & Year(d)
End Function

Function FormatDateShort(ByVal d)
    If IsNull(d) Or Not IsDate(d) Then
        FormatDateShort = "-"
        Exit Function
    End If
    FormatDateShort = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d)
End Function

Function FormatDateInput(ByVal d)
    If IsNull(d) Or Not IsDate(d) Then
        FormatDateInput = ""
        Exit Function
    End If
    FormatDateInput = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
End Function

' Lista archivos de una carpeta virtual como texto separado por "|"
' (se usa para los selectores de imágenes y archivos de audio del panel admin)
Function ListFilesCsv(ByVal virtualFolder, ByVal extensions)
    Dim folder, f, ext, i, ok, out
    out = ""
    On Error Resume Next
    Set folder = Fso().GetFolder(Server.MapPath(AppBase() & virtualFolder))
    If Err.Number = 0 Then
        For Each f In folder.Files
            ext = LCase(Fso().GetExtensionName(f.Name))
            ok = False
            For i = 0 To UBound(extensions)
                If ext = extensions(i) Then ok = True
            Next
            If ok Then
                If out <> "" Then out = out & "|"
                out = out & "~/" & virtualFolder & "/" & f.Name
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
    ListFilesCsv = out
End Function

' ---------------------------------------------------------------------
'  Presentacion: portadas, avatares, iniciales
' ---------------------------------------------------------------------
Function Initials(ByVal label)
    Dim s, parts, i, out
    s = Trim(CStr(label & ""))
    If s = "" Then
        Initials = "?"
        Exit Function
    End If
    parts = Split(Replace(s, "-", " "), " ")
    out = ""
    For i = 0 To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then
            out = out & UCase(Left(Trim(parts(i)), 1))
            If Len(out) >= 2 Then Exit For
        End If
    Next
    If out = "" Then out = UCase(Left(s, 1))
    Initials = out
End Function

Sub PaletteFor(ByVal label, ByRef c1, ByRef c2)
    Dim i, h, p1, p2
    h = 0
    For i = 1 To Len(label & "")
        h = h + AscW(Mid(label & "", i, 1))
    Next
    p1 = Array("#38bdf8", "#3b82f6", "#22d3ee", "#0ea5e9", "#34d399", "#2dd4bf")
    p2 = Array("#34d399", "#22d3ee", "#6ee7b7", "#34d399", "#38bdf8", "#3b82f6")
    c1 = p1(Abs(h) Mod 6)
    c2 = p2(Abs(h) Mod 6)
End Sub

' Portada cuadrada con respaldo generativo si el archivo no existe
Function Cover(ByVal imgPath, ByVal label, ByVal extraClass)
    Dim c1, c2
    If AssetExists(imgPath) Then
        Cover = "<span class=""cover " & extraClass & """><img src=""" & HtmlAttr(AssetUrl(imgPath)) & _
                """ alt=""" & HtmlAttr(label) & """ loading=""lazy"" decoding=""async""></span>"
    Else
        PaletteFor label, c1, c2
        Cover = "<span class=""cover cover--fb " & extraClass & """ style=""--c1:" & c1 & ";--c2:" & c2 & """>" & _
                "<i>" & HtmlEncode(Initials(label)) & "</i></span>"
    End If
End Function

' Avatar circular con respaldo generativo
Function Avatar(ByVal imgPath, ByVal label, ByVal extraClass)
    Dim c1, c2
    If AssetExists(imgPath) Then
        Avatar = "<span class=""avatar " & extraClass & """><img src=""" & HtmlAttr(AssetUrl(imgPath)) & _
                 """ alt=""" & HtmlAttr(label) & """ loading=""lazy""></span>"
    Else
        PaletteFor label, c1, c2
        Avatar = "<span class=""avatar avatar--fb " & extraClass & """ style=""--c1:" & c1 & ";--c2:" & c2 & """>" & _
                 "<i>" & HtmlEncode(Initials(label)) & "</i></span>"
    End If
End Function

' ---------------------------------------------------------------------
'  Mensajes flash y CSRF
' ---------------------------------------------------------------------
Sub Flash(ByVal msg, ByVal kind)
    Session("flash_msg") = msg
    Session("flash_kind") = kind
End Sub

Function TakeFlash()
    Dim f
    f = Session("flash_msg") & ""
    If f <> "" Then
        TakeFlash = (Session("flash_kind") & "|") & f
    Else
        TakeFlash = ""
    End If
    Session("flash_msg") = ""
    Session("flash_kind") = ""
End Function

Function CsrfToken()
    Dim t
    t = Session("csrf_token") & ""
    If t = "" Then
        Randomize
        t = Hex(CLng(Timer * 997)) & Hex(CLng(Rnd() * 1000000)) & Hex(CLng(Rnd() * 1000000))
        Session("csrf_token") = t
    End If
    CsrfToken = t
End Function

Function CsrfField()
    CsrfField = "<input type=""hidden"" name=""csrf"" value=""" & CsrfToken() & """>"
End Function

Function CsrfValid()
    Dim t
    t = Request.Form("csrf") & ""
    If t = "" Then t = Request.ServerVariables("HTTP_X_CSRF_TOKEN") & ""
    CsrfValid = (t <> "" And t = (Session("csrf_token") & ""))
End Function

Sub CsrfRequire()
    If Not CsrfValid() Then
        Response.Status = "403 Forbidden"
        Response.ContentType = "text/plain"
        Response.Write "Solicitud rechazada: token de seguridad invalido."
        Response.End
    End If
End Sub

' ---------------------------------------------------------------------
'  Redirecciones seguras
' ---------------------------------------------------------------------
Function CurrentPath()
    Dim u
    u = Request.ServerVariables("URL") & ""
    If Request.ServerVariables("QUERY_STRING") & "" <> "" Then
        u = u & "?" & Request.ServerVariables("QUERY_STRING")
    End If
    CurrentPath = u
End Function

' Solo se permite volver a una ruta interna de la aplicación
Function SafeNext(ByVal n)
    Dim s
    s = Trim(n & "")
    If s = "" Then
        SafeNext = ""
        Exit Function
    End If
    s = Replace(s, "\", "/")
    If Left(s, 1) <> "/" Then s = "/" & s
    If InStr(s, "//") > 0 Or InStr(s, ":") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        SafeNext = ""
        Exit Function
    End If
    If Left(s, Len(AppBase())) <> AppBase() Then
        SafeNext = ""
        Exit Function
    End If
    SafeNext = s
End Function

Sub GoToPage(ByVal url)
    Response.Redirect url
    Response.End
End Sub

' ---------------------------------------------------------------------
'  Iconografia SVG (trazo 1.7, 24x24)
' ---------------------------------------------------------------------
Function Icon(ByVal name, ByVal cls)
    Dim open, close, body
    open = "<svg class=""ico " & cls & """ viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" " & _
           "stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">"
    close = "</svg>"
    Select Case LCase(name)
        Case "home"     : body = "<path d=""M3 10.6 12 3l9 7.6""/><path d=""M5.2 9.4V21h13.6V9.4""/><path d=""M9.6 21v-6.2h4.8V21""/>"
        Case "compass"  : body = "<circle cx=""12"" cy=""12"" r=""9""/><path d=""m15.6 8.4-2.3 5.1-5.1 2.3 2.3-5.1z""/>"
        Case "music"    : body = "<path d=""M9 17.6V6.2l10-2.1v11.4""/><circle cx=""6.4"" cy=""18"" r=""2.6""/><circle cx=""16.4"" cy=""15.9"" r=""2.6""/>"
        Case "mic"      : body = "<rect x=""9"" y=""3"" width=""6"" height=""10"" rx=""3""/><path d=""M5 11a7 7 0 0 0 14 0""/><path d=""M12 18v3""/>"
        Case "disc"     : body = "<circle cx=""12"" cy=""12"" r=""9""/><circle cx=""12"" cy=""12"" r=""2.4""/>"
        Case "list"     : body = "<path d=""M8.5 6h12M8.5 12h12M8.5 18h12""/><path d=""M3.6 6h.01M3.6 12h.01M3.6 18h.01""/>"
        Case "users"    : body = "<circle cx=""9"" cy=""8"" r=""3.2""/><path d=""M2.8 20a6.2 6.2 0 0 1 12.4 0""/><path d=""M16 5.6a3 3 0 0 1 0 5.6""/><path d=""M17.4 14.4A6.2 6.2 0 0 1 21.2 20""/>"
        Case "shield"   : body = "<path d=""M12 3.2 19 6v5.4c0 4.3-2.9 7.6-7 9.4-4.1-1.8-7-5.1-7-9.4V6z""/><path d=""m9 12 2.1 2.1L15 10.2""/>"
        Case "logout"   : body = "<path d=""M15 4.5h3A2 2 0 0 1 20 6.5v11a2 2 0 0 1-2 2h-3""/><path d=""M10.5 8.2 6.7 12l3.8 3.8""/><path d=""M6.7 12H16""/>"
        Case "search"   : body = "<circle cx=""11"" cy=""11"" r=""6.8""/><path d=""m20 20-3.6-3.6""/>"
        Case "play"     : body = "<path d=""M8.2 5.4v13.2L19 12z"" fill=""currentColor"" stroke=""none""/>"
        Case "pause"    : body = "<rect x=""7"" y=""5"" width=""3.4"" height=""14"" rx=""1.3"" fill=""currentColor"" stroke=""none""/><rect x=""13.6"" y=""5"" width=""3.4"" height=""14"" rx=""1.3"" fill=""currentColor"" stroke=""none""/>"
        Case "prev"     : body = "<path d=""M18.5 6.2v11.6L9.6 12z"" fill=""currentColor"" stroke=""none""/><rect x=""5.4"" y=""6.2"" width=""2.2"" height=""11.6"" rx=""1.1"" fill=""currentColor"" stroke=""none""/>"
        Case "next"     : body = "<path d=""M5.5 6.2v11.6L14.4 12z"" fill=""currentColor"" stroke=""none""/><rect x=""16.4"" y=""6.2"" width=""2.2"" height=""11.6"" rx=""1.1"" fill=""currentColor"" stroke=""none""/>"
        Case "volume"   : body = "<path d=""M4 9.4h3.2L11.4 6v12L7.2 14.6H4z""/><path d=""M14.8 9.2a4 4 0 0 1 0 5.6""/><path d=""M17.4 6.6a7.6 7.6 0 0 1 0 10.8""/>"
        Case "volume-x" : body = "<path d=""M4 9.4h3.2L11.4 6v12L7.2 14.6H4z""/><path d=""m15.4 9.6 4.4 4.8M19.8 9.6l-4.4 4.8""/>"
        Case "plus"     : body = "<path d=""M12 5.2v13.6M5.2 12h13.6""/>"
        Case "minus"    : body = "<path d=""M5.2 12h13.6""/>"
        Case "trash"    : body = "<path d=""M4 7h16""/><path d=""M9.5 7V4.6h5V7""/><path d=""M6.4 7 7.4 20h9.2l1-13""/><path d=""M10.4 11v5.2M13.6 11v5.2""/>"
        Case "edit"     : body = "<path d=""M4 20h4.2L20 8.2 15.8 4 4 15.8z""/><path d=""m14.4 5.4 4.2 4.2""/>"
        Case "close"    : body = "<path d=""m6.2 6.2 11.6 11.6M17.8 6.2 6.2 17.8""/>"
        Case "chevron-down" : body = "<path d=""m6 9.6 6 6 6-6""/>"
        Case "chevron-right": body = "<path d=""m9.6 6 6 6-6 6""/>"
        Case "arrow-left"   : body = "<path d=""M19 12H5.4""/><path d=""m11 5.8-6 6.2 6 6.2""/>"
        Case "menu"     : body = "<path d=""M4 7h16M4 12h16M4 17h16""/>"
        Case "queue"    : body = "<path d=""M4 7h10.5M4 12h10.5M4 17h7""/><path d=""m16.4 15.2 4.4 2.9-4.4 2.9z"" fill=""currentColor"" stroke=""none""/>"
        Case "shuffle"  : body = "<path d=""M3 6.6h3.6l9.8 10.8H21""/><path d=""M3 17.4h3.6l3.2-3.5""/><path d=""M14.6 6.6H21""/><path d=""m18.2 3.4 3 3.2-3 3.2""/><path d=""m18.2 14.2 3 3.2-3 3.2""/>"
        Case "repeat"   : body = "<path d=""M4.2 9.2V8a3.2 3.2 0 0 1 3.2-3.2h11""/><path d=""m15.6 2 3 2.8-3 2.8""/><path d=""M19.8 14.8V16a3.2 3.2 0 0 1-3.2 3.2h-11""/><path d=""m8.4 22-3-2.8 3-2.8""/>"
        Case "check"    : body = "<path d=""m5 12.6 4.6 4.6L19 7.4""/>"
        Case "alert"    : body = "<path d=""M12 4.2 2.9 20h18.2z""/><path d=""M12 9.6v4.4M12 17h.01""/>"
        Case "info"     : body = "<circle cx=""12"" cy=""12"" r=""9""/><path d=""M12 11v5.2M12 7.8h.01""/>"
        Case "mail"     : body = "<rect x=""3"" y=""5.2"" width=""18"" height=""13.6"" rx=""2.6""/><path d=""m4 7.4 8 5.8 8-5.8""/>"
        Case "calendar" : body = "<rect x=""3.5"" y=""5.2"" width=""17"" height=""15.3"" rx=""2.6""/><path d=""M8 3v4.2M16 3v4.2M3.5 10.2h17""/>"
        Case "user"     : body = "<circle cx=""12"" cy=""8"" r=""3.6""/><path d=""M4.6 20.6a7.4 7.4 0 0 1 14.8 0""/>"
        Case "camera"   : body = "<path d=""M4 8.6h3.1L8.6 6.4h6.8l1.5 2.2H20a1.6 1.6 0 0 1 1.6 1.6v8A1.6 1.6 0 0 1 20 19.8H4a1.6 1.6 0 0 1-1.6-1.6v-8A1.6 1.6 0 0 1 4 8.6z""/><circle cx=""12"" cy=""13.8"" r=""3.2""/>"
        Case "folder"   : body = "<path d=""M3 7.4A2 2 0 0 1 5 5.4h4l2 2.6h8a2 2 0 0 1 2 2v7.4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z""/>"
        Case "grid"     : body = "<rect x=""3.6"" y=""3.6"" width=""7"" height=""7"" rx=""2.2""/><rect x=""13.4"" y=""3.6"" width=""7"" height=""7"" rx=""2.2""/><rect x=""3.6"" y=""13.4"" width=""7"" height=""7"" rx=""2.2""/><rect x=""13.4"" y=""13.4"" width=""7"" height=""7"" rx=""2.2""/>"
        Case "sparkles" : body = "<path d=""M11.4 3.6 13.2 9l5.4 1.8-5.4 1.8-1.8 5.4-1.8-5.4L4.2 10.8 9.6 9z""/><path d=""M18.4 16.2l.7 2.1 2.1.7-2.1.7-.7 2.1-.7-2.1-2.1-.7 2.1-.7z""/>"
        Case "lock"     : body = "<rect x=""4.6"" y=""10.4"" width=""14.8"" height=""10.2"" rx=""2.6""/><path d=""M8 10.4V8a4 4 0 0 1 8 0v2.4""/>"
        Case "clock"    : body = "<circle cx=""12"" cy=""12"" r=""9""/><path d=""M12 7.4V12l3.1 2""/>"
        Case "save"     : body = "<path d=""M5 4h10.8L19 7.2V20H5z""/><path d=""M8.2 4v5h6.6V4""/><rect x=""8.2"" y=""13.4"" width=""7.6"" height=""6.6"" rx=""1""/>"
        Case "more"     : body = "<circle cx=""5.4"" cy=""12"" r=""1.5"" fill=""currentColor"" stroke=""none""/><circle cx=""12"" cy=""12"" r=""1.5"" fill=""currentColor"" stroke=""none""/><circle cx=""18.6"" cy=""12"" r=""1.5"" fill=""currentColor"" stroke=""none""/>"
        Case "headphones": body = "<path d=""M4 15v-2.6a8 8 0 0 1 16 0V15""/><rect x=""2.8"" y=""14"" width=""4.4"" height=""6.6"" rx=""2.2""/><rect x=""16.8"" y=""14"" width=""4.4"" height=""6.6"" rx=""2.2""/>"
        Case "wave"     : body = "<path d=""M3 12h2.4l1.8-5 2.4 10 2.4-14 2.4 16 2.4-9 1.8 2H21""/>"
        Case "eye"      : body = "<path d=""M2.6 12S6 6.4 12 6.4 21.4 12 21.4 12 18 17.6 12 17.6 2.6 12 2.6 12z""/><circle cx=""12"" cy=""12"" r=""2.9""/>"
        Case "eye-off"  : body = "<path d=""M4 4.4 20 19.6""/><path d=""M9.6 7.1A9.7 9.7 0 0 1 12 6.8c6 0 9.4 5.2 9.4 5.2a16 16 0 0 1-2.8 3.3""/><path d=""M6.5 8.4A15.9 15.9 0 0 0 2.6 12s3.4 5.2 9.4 5.2a9.6 9.6 0 0 0 3.6-.7""/><path d=""M10.2 10.4a2.6 2.6 0 0 0 3.5 3.6""/>"
        Case Else       : body = "<circle cx=""12"" cy=""12"" r=""8.4""/>"
    End Select
    Icon = open & body & close
End Function

' ---------------------------------------------------------------------
'  Capa de acceso a datos (ADO + consultas parametrizadas)
' ---------------------------------------------------------------------
Function Db()
    If gConn Is Nothing Then
        Set gConn = Server.CreateObject("ADODB.Connection")
        gConn.ConnectionTimeout = 15
        gConn.CommandTimeout = 30
        gConn.Open DB_CONN_STRING
    ElseIf gConn.State = 0 Then
        gConn.Open DB_CONN_STRING
    End If
    Set Db = gConn
End Function

Sub DbClose()
    On Error Resume Next
    If Not (gConn Is Nothing) Then
        If gConn.State <> 0 Then gConn.Close
    End If
    Set gConn = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Function BuildParam(ByVal cmd, ByVal v)
    Dim p
    If IsNull(v) Then
        Set p = cmd.CreateParameter("p", adVarWChar, adParamInput, 1, Null)
    ElseIf VarType(v) = vbBoolean Then
        Set p = cmd.CreateParameter("p", adBoolean, adParamInput, 1, v)
    ElseIf VarType(v) = vbDate Then
        Set p = cmd.CreateParameter("p", adDBTimeStamp, adParamInput, 16, v)
    ElseIf IsNumeric(v) And VarType(v) <> vbString Then
        Set p = cmd.CreateParameter("p", adInteger, adParamInput, 4, CLng(v))
    Else
        Set p = cmd.CreateParameter("p", adVarWChar, adParamInput, 4000, CStr(v))
    End If
    Set BuildParam = p
End Function

' Devuelve un Recordset desconectado (cursor estatico en cliente)
Function DbQuery(ByVal sql, ByVal params)
    Dim cmd, rs, i
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = Db()
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    cmd.CommandTimeout = 30
    If IsArray(params) Then
        For i = 0 To UBound(params)
            cmd.Parameters.Append BuildParam(cmd, params(i))
        Next
    End If
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    Set rs.ActiveConnection = Nothing
    Set DbQuery = rs
End Function

Function DbExecute(ByVal sql, ByVal params)
    Dim cmd, i, n
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = Db()
    cmd.CommandType = adCmdText
    cmd.CommandText = sql
    cmd.CommandTimeout = 30
    If IsArray(params) Then
        For i = 0 To UBound(params)
            cmd.Parameters.Append BuildParam(cmd, params(i))
        Next
    End If
    n = 0
    cmd.Execute n, , adExecuteNoRecords
    DbExecute = n
End Function

' Primer campo de la primera fila
Function DbScalar(ByVal sql, ByVal params)
    Dim rs
    Set rs = DbQuery(sql, params)
    If rs.EOF Then
        DbScalar = Null
    Else
        DbScalar = rs.Fields(0).Value
    End If
    rs.Close
    Set rs = Nothing
End Function
%>
