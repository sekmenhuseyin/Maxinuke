<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%call ORTA%>
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												   Modul Uygulamasi													#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
module_id = Request.QueryString("mi")
module_id = QS_CLEAR(module_id)
If module_id = "" Then
response.redirect "./."
End if
set modulgel = server.createobject("adodb.recordset")
modulgelsql = "select * from moduller where module_id="&module_id&""
modulgel.open modulgelsql, Connection, 1, 3
IF modulgel.Eof OR modulgel.Bof Then
%>
<table border="0" cellpadding="2" cellspacing="3" width="100%" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td align="center" bgcolor="<%=content_bg%>"><BR><BR><BR><b><%=module_eof%></b><BR><BR><BR><BR></td>
	</tr>
</table>
<%ELSE%>
<table border="0" cellpadding="2" cellspacing="3" width="100%" style="<%=TableBorder%>" class="td_menu" bgcolor="<%=content_bg%>">
	<tr>
		<td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=modulgel("mname")%></B></CENTER></td>
	</tr>
	<tr>
		<td>
<%
	If modulgel("mems") = True Then
		If Session("Enter") = "Yes" Then
		Server.Execute ("moduller\"&modulgel("mdir")&"\default.asp")
		Else
		Response.Write sett("np_msg")
		End If
	ELSE
	Server.Execute ("moduller\"&modulgel("mdir")&"\default.asp")
	End IF
%>
<!--#include file="inc/adsense_3.asp" -->
		</td>
	</tr>
</table>
<%
End IF
modulgel.close
set modulgel=nothing
call ORTA2
call ALT
%>