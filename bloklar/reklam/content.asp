<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												Reklam Blok Uygulamasi												#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
set minireklam = server.createobject("adodb.recordset")
SQL = "Select * FROM BANNERS WHERE active = True AND position = 'yan'"
minireklam.open SQL,Connection,1,3
%>
<table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
for k=1 to minireklam.RecordCount
if minireklam.eof or minireklam.bof then exit for
%>
	<TR>
		<TD align="center">
<%IF minireklam("ban_type") = "Flash" Then%>
		<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" WIDTH="135" HEIGHT="31" onClick="javascript:location.href('redirect.asp?go=ads&bid=<%=minireklam("banner_id")%>');">
		<PARAM NAME="movie" VALUE="<%=minireklam("ban_url")%>">
		<param name="WMode" value="Transparent">
		<PARAM NAME="quality" VALUE="high">
		<EMBED src="<%=minireklam("ban_url")%>" quality="high" WIDTH="135" HEIGHT="31"TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED></OBJECT>
		<SCRIPT src="inc/mn.js" type=text/javascript></SCRIPT>
<% ELSEIF minireklam("ban_type") = "Normal" Then %>
		<a href="redirect.asp?go=ads&bid=<%=minireklam("banner_id")%>" TARGET="_new"><img src="<%=minireklam("ban_url")%>" width="135" border="0" align="absmiddle"></a>
<% ELSEIF minireklam("ban_type") = "Kod" Then %>
		<%=minireklam("ban_url")%>
		<%End IF%>
		</TD>
	</TR>
<%minireklam.movenext : next%>
</TABLE>
<% 
minireklam.close
set minireklam=nothing
Connection.close
set Connection=nothing
%>