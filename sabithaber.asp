<%
set sabit_haber = server.createobject("adodb.recordset")
SQL = "select * from FIXED order by fid desc"
sabit_haber.open SQL, Connection, 1, 3
if sabit_haber.eof then
else
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=content_bg%>" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td height="25" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><font size="3"><b><%=sabit_haber("sf_baslik")%></b></font></td>
	</tr>
	<tr>
		<td style="padding-left:10px;padding-right:10px"><%=sabit_haber("f_metin")%></td>
	</tr>
</table>
<%
End if
sabit_haber.close
set sabit_haber=nothing
%>