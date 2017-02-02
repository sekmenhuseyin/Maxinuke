<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Then
	IF Request.QueryString("op") = "New" Then
%>
<form action="?op=Register" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td align="right">Tema Adý : </td>
		<td align="left"><input type="text" name="theme_name" class="form1"></td>
	</tr>
	<tr>
		<td align="right">Tema Klasörü : </td>
		<td align="left">images/temalar/<input type="text" name="theme_dir" class="form1"></td>
	</tr>
	<tr>
		<td align="CEnter" colspan=2><input type="submit" value="Temayý Ekle »" class="submit"></td>
	</tr>
</table>
</form>
<%
	ElseIF Request.QueryString("op") = "Register" Then
name = Request.Form("theme_name")
dir = Request.Form("theme_dir")
theme_status = "True"

IF name = "" OR dir = "" Then
Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz...</b></div>"
ELSE
Connection.Execute("INSERT INTO THEMES (t_name,t_dir,t_active) VALUES ('"&name&"','"&dir&"',"&theme_status&")")
Response.Redirect "themes.asp"
End IF
	ElseIF Request.QueryString("op") = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM THEMES WHERE t_id = "&Request.QueryString("id")&"",Connection,3,3
%>
<form action="?op=Update&id=<%=Request.QueryString("id")%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td align="right">Tema Adý : </td>
		<td align="left"><input type="text" name="theme_name" class="form1"  value="<%=rs("t_name")%>"></td>
	</tr>
	<tr>
		<td align="right">Tema Klasörü : </td>
		<td align="left">images/temalar/<input type="text" name="theme_dir" class="form1" value="<%=rs("t_dir")%>"></td>
	</tr>
	<tr>
		<td align="center" colspan=2><input type="submit" value="Güncelle »" class="submit"></td>
	</tr>
</table>
</form>
<%
rs.Close
Set rs = Nothing
	ElseIF Request.QueryString("op") = "Update" Then

t_name = Request.Form("theme_name")
t_dir = Request.Form("theme_dir")


IF t_name = "" OR t_dir = "" Then
Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz...</b></div>"
ELSE
Connection.Execute("UPDATE THEMES SET t_name = '"&t_name&"' WHERE t_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE THEMES SET t_dir = '"&t_dir&"' WHERE t_id = "&Request.QueryString("id")&"")
Response.Redirect "themes.asp"
End IF

	ElseIF Request.QueryString("op") = "Delete" Then
'		eskita=Request.QueryString("ta")
'		Set temalasifirla=Server.CreateObject("Adodb.Recordset")
'		temalasifirla.ActiveConnection = Connection
'		Sorgu="SELECT * FROM MEMBERS WHERE u_theme="&eskita&""
'		temalasifirla.open Sorgu,Connection,3,2
'		temalasifirla.update "u_theme",standart
'		temalasifirla.close
'		set temalasifirla=nothing
		Connection.Execute("DELETE FROM THEMES WHERE t_id = "&Request.QueryString("id")&"")
		Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ElseIF Request.QueryString("op") = "Change" Then
		Connection.Execute("UPDATE THEMES SET t_active = "&Request.QueryString("isl")&" WHERE t_id = "&Request.QueryString("id")&"")
		Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM THEMES ORDER BY t_id",Connection,3,3
no=0
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><div class="block_title"><B>-=- TEMALAR -=-</B></div></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">No</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tema Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tema Klasörü</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Durum</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
IF rs("t_active") = True Then
islem = "Pasifleþtir"
status = "Aktif"
isl = "False"
ELSe
islem = "Aktifleþtir"
status = "Pasif"
isl = "True"
End IF
no=no+1
%>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td align="center"><%=no%></td>
		<td><%=rs("t_name")%></td>
		<td align="center"><%=rs("t_dir")%></td>
		<td align="center"><%=status%></td>
		<td align="center"><a href="?op=Change&isl=<%=isl%>&id=<%=rs("t_id")%>"><%=islem%></a> - <a href="?op=Edit&id=<%=rs("t_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("t_id")%>&ta=<%=rs("t_name")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
	End IF
End IF
%>