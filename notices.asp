<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%
IF Session("Level") = "1" Or Session("Level") = "4" Then
action = Request.QueryString("x")

IF action = "New" Then
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu" align="center">
<form action="?x=Register" method="post">
<tr>
<td colspan="2"><textarea name="notice_text" rows="10" class="form1" style="width:100%"></textarea></td>
</tr>
<tr>
<td width="50%" align="center" colspan="2"><input type="submit" value="     DUYURUYU EKLE     " class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF action = "Register" Then
n_text = Request.Form("notice_text")
IF n_text = "" Then
Response.Write "<div align=center>Lütfen Duyuruyu Yazýnýz</div>"
ELSE
Set nRs = Server.CreateObject("ADODB.RecordSet")
nRs.open "Select * FROM NOTICES",Connection,3,3

nRs.AddNew
nRs("n_text") = n_text
nRs("n_date") = Now()
nRs.Update
Response.Redirect "notices.asp"
End IF
ElseIF action = "Edit" Then
Set rs = Connection.Execute("Select * FROM NOTICES WHERE n_id = "&Request.QueryString("id")&"")
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu" align="center">
<form action="?x=Update&id=<%=Request.QueryString("id")%>" method="post">
<tr>
<td><textarea name="notice_text" rows="10" style="width:100%" class="form1"><%=rs("n_text")%></textarea></td>
</tr>
<tr>
<td align="center"><input type="submit" value="     DUYURUYU DUZENLE     " class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF action = "Update" Then
n_text = Request.Form("notice_text")
IF n_text = "" Then
Response.Write "<div align=center>Lütfen Duyuruyu Yazýnýz</div>"
ELSE
Set uRs = Server.CreateObject("ADODB.RecordSet")
uRs.open "Select * FROM NOTICES WHERE n_id = "&Request.QueryString("id")&"",Connection,3,3
uRs("n_text") = n_text
uRs.Update
Response.Redirect "notices.asp"
End IF
ElseIF action = "Delete" Then
Set d_n = Connection.Execute("DELETE FROM NOTICES WHERE n_id = "&Request.QueryString("id")&"")
Set d_n = Nothing
Response.Redirect "notices.asp"
ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NOTICES",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><B>-=- DUYURULAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">No</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Duyuru (Ýlk 150 harf)</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% Do Until rs.Eof %>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td align="center"><%=rs("n_id")%></td>
		<td align="center"><%=Left(rs("n_text"),150)%></td>
		<td align="center"><a href="?x=Edit&id=<%=rs("n_id")%>">Düzenle</a> - <a href="?x=Delete&id=<%=rs("n_id")%>">Sil</a></td>
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