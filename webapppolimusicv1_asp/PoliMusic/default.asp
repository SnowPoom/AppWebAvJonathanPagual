<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Punto de entrada del sitio: redirige segun el estado de la sesión
If IsLoggedIn() Then
    GoToPage AppBase() & "home.asp"
Else
    GoToPage AppBase() & "login.asp"
End If
%>
