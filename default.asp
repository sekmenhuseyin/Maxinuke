<%Response.Buffer = True%>
<!--#include file="anakontrol.asp" -->
<!--#include file="inc/includes.asp" -->
<%

If session("radyo")="" then
	If sett("rplayer") = True then
	session("radyo") = "gorun"
	else
	session("radyo") = "gizlen"
	End if
End if
If session("radyo")="gizlen" Then
response.redirect "defaultt.asp"
else
%>
<html>
<head>
<title><%=sett("site_isim")%> - <%=sett("site_slogan")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<LINK href="images/site/favicon.ico" rel="SHORTCUT ICON">
<META NAME="TITLE" CONTENT="<%=sett("site_isim")%> - <%=sett("site_slogan")%>">
<meta http-equiv="Content-Language" content="tr">
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="Maxinuke">
<META NAME="Description" CONTENT="<%=sett("Keywords")%>">
<META NAME="Keywords" CONTENT="<%=sett("Keywords")%>,Maxinuke">
<META NAME="Subject" CONTENT="<%=sett("Keywords")%>">
<meta name="abstract" content="<%=sett("site_isim")%>">
<META NAME="Copyright" CONTENT="Maxinuke © 2008 - 2009">
<META NAME="Distribution" CONTENT="Global">
<meta name="robots" content="index, follow, imageindex, imageclick, archive, all" >
<meta NAME="rating" content="All">
<meta name="publisher" content="<%=sett("site_adres")%>">
<meta name="audience" content="all">
<meta name="REVISIT-AFTER" content="1 days">
<meta http-equiv="reply-to" content="<%=sett("site_mail")%>">
<meta name="resource-type" content="document">
<META NAME="googlebot" CONTENT="Index, Follow">
<meta NAME="Designer" CONTENT="Maxinuke Team and Penoaydin"/>
<META HTTP-EQUIV="imagetoolbar" CONTENT="no">
</head>
<style>
frameset {
margin:0px;
padding:0px;
border:0px;
background-color:buttonface;
}
</style>
<frameset rows="*,8" rows1="*,30" name="fst">
	<frame src="defaultt.asp" frameborder="0">
	<frame src="player/player.asp" frameborder="0">
</frameset>
<%End if%>