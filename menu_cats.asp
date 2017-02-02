<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Then
	IF Request.QueryString("op") = "New" Then
	Set toplamkat = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * from MENU_CATS"
	toplamkat.Open SQL,Connection,1,3
	toplami=toplamkat.recordcount
	%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<form action="?op=Register" method="post">
	<tr>
		<td colspan="5" align="center"><B>-=- YENI MENU KATEGORISI EKLEME -=-</B></td>
	</tr>
	<tr>
		<td align="center">
		Kategori Adý<BR>
		<input type="text" name="title" class="form1"><BR><BR>
		Dizilim Sýrasý<BR>
		<select name="menu_queue" size="1" class="form1">
		<%
		For I = 1 TO toplami+1%>
		<option value="<%=I%>"><%=I%></option>
		<% Next %>
		</select>
<%
toplamkat.close
set toplamkat=nothing
%>
		<BR><BR>
		<input type="submit" value="Kategoriyi Ekle »" class="submit"></td>
	</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op") = "Register" Then
cat_title = duz(Request.Form("title"))
IF cat_title = "" Then
Response.Write "<div align=center class=td_menu>Lütfen Kategori Adýný yazýnýz...</div>"
ELSE
Set ne = Server.CreateObject("ADODB.RecordSet")
ne.open "Select * FROM MENU_CATS",Connection,3,3
ne.AddNew
ne("mc_title") = cat_title
ne("mc_queue") = Request.Form("menu_queue")
ne.Update
ne.Close
Set ne = Nothing
Response.Redirect "menu_cats.asp"
End IF
	ElseIF Request.QueryString("op")  = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MENU_CATS WHERE mc_id = "&Request.QueryString("id")&"",Connection,3,3
	Set toplamkat = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * from MENU_CATS"
	toplamkat.Open SQL,Connection,1,3
	toplami=toplamkat.recordcount
%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="td_menu">
<tr><td colspan="2" ></td></tr>
<form action="?op=Update&id=<%=Request.QueryString("id")%>" method="post">
<tr>
<td width="35%" align="right">Kategori Adý : </td>
<td width="65%" align="left"><input type="text" name="title" size="50" class="form1" value="<%=rs("mc_title")%>"></td>
</tr>
<tr >
<td width="35%" align="right">Sýra : </td>
<td width="65%" align="left">
<select name="queue" size="1" class="form1">
<%
For I = 1 TO toplami+1
IF I = rs("mc_queue") Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<% Next %>
</select>
<%
toplamkat.close
set toplamkat=nothing
%>
</td>
</tr>
<tr >
<td width="35%" align="right"></td>
<td width="65%" align="left"><input type="submit" value="Güncelle »" class="submit"></td>
</tr>
</form>
<tr><td colspan="2" ></td></tr>
</table>
<%
rs.Close
Set rs = Nothing
	ElseIF Request.QueryString("op") = "Update" Then
cat_title = duz(Request.Form("title"))
IF cat_title = "" Then
Response.Write "<div align=center class=td_menu>Lütfen Kategori Adýný yazýnýz...</div>"
ELSE
Connection.Execute("UPDATE MENU_CATS Set mc_title = '"&cat_title&"' WHERE mc_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE MENU_CATS Set mc_queue = '"&Request.Form("queue")&"' WHERE mc_id = "&Request.QueryString("id")&"")
Response.Redirect "menu_cats.asp"
End IF
	ElseIF Request.QueryString("op") = "Delete" Then
Connection.Execute("DELETE FROM MENU_CATS WHERE mc_id ="&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MENU_CATS ORDER BY mc_queue",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><B>-=- MENU KATEGORILERI -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Sýrasý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% Do Until rs.EoF %>
<tr onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
<td align="center"><%=rs("mc_title")%></td>
<td align="center"><%=rs("mc_queue")%></td>
<td align="center"><a href="?op=Edit&id=<%=rs("mc_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("mc_id")%>">Sil</a></td>
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