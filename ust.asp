<%Response.Buffer = True%>
<!--#include file="anakontrol.asp" -->
<!--#include file="inc/includes.asp" -->
<!--#include file="bh.asp" -->
<html>
<head>
<title><%=sett("site_isim")%> - <%=sett("site_slogan")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<LINK href="images/site/favicon.ico" rel="SHORTCUT ICON">
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<META NAME="TITLE" CONTENT="<%=sett("site_isim")%> - <%=sett("site_slogan")%>">
<meta http-equiv="Content-Language" content="tr">
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
<body bgcolor="<%=background%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
</head>
<!--#include file="duyuru.asp" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="<%=background%>">
  <tr> 
    <td>
	<%if sett("logo_tipi")=False then%>
	<A HREF="defaultt.asp"><IMG SRC="<%=sett("site_logo")%>" BORDER="0" ALT="<%=sett("site_isim")%>" align="absmiddle"></A>
	<%elseif sett("logo_tipi")=true then%>
	<OBJECT codeBase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,115,0" height="60" width="310" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000">
	<PARAM NAME="movie" VALUE="<%=sett("site_logo")%>">
	<param name="WMode" value="Transparent">
	<PARAM NAME="quality" VALUE="high">
	<EMBED src="<%=sett("site_logo")%>" quality="high" WIDTH="310" HEIGHT="60"TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED></OBJECT>
	<SCRIPT src="inc/mn.js" type=text/javascript></SCRIPT>
	<%End if%>
	</td>
    <td align="right"><!--#include file="reklam_ust.asp" --></td>
  </tr>
</table>
<!--#include file="ust_linkler.asp" -->
<!--#include file="uyepanel.asp" -->