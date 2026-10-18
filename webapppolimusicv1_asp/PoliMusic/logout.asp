<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/functions.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' Cierre de sesión: destruye la sesión de ASP Clasico y vuelve al login
Dim back
back = AppBase() & "login.asp"
EndUserSession()
Response.Redirect back
%>
