<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Then
	IF Request.QueryString("action")= "" Then
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open "Select * FROM SWEARWORDS",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><div class="block_title">-=- ARGO KELÝME LÝSTESÝ -=-</div></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Argo Kelime</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Gösterilen Deðer</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%Do Until rs.EoF%>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td align="center"><%=rs("id")%></td>
		<td><%=rs("s_text")%></td>
		<td><%=rs("s_value")%></td>
		<td align="center"><a href="?action=Edit&id=<%=rs("id")%>">Düzenle</a> - <a href="?action=Delete&id=<%=rs("id")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
elseIF Request.QueryString("action")= "Edit" Then
Set rs = Connection.Execute("Select * FROM SWEARWORDS WHERE id = "&Request.QueryString("id")&"")
%>
<form action="?action=Update&id=<%=Request.QueryString("id")%>" method="post">
<table border="0" cellspacing="2" cellpading="2" width="75%" align="center" class="td_menu">
<tr >
<td width="40%" align="right">Argo Kelime : </td>
<td width="60%"><input type="text" name="s_text" class="form1"  value="<%=rs("s_text")%>"></td>
</tr>
<tr >
<td width="40%" align="right">Görüntülenen Deðer : </td>
<td width="60%"><input type="text" name="s_value" class="form1"  value="<%=rs("s_value")%>"></td>
</tr>
<tr >
<td align="center" colspan=2><input type="submit" value="Güncelle" class="submit"></td>
</tr>
</table>
</form>
<%
elseIF Request.QueryString("action")= "Update" Then
s_text = duz(Request.Form("s_text"))
s_value = duz(Request.Form("s_value"))
			IF s_text = "" OR s_value = "" Then
			Response.Write "<div align=center class=td_menu>Tüm Alanlarý Doldurunuz !</div>"
			ELSE
				SET upd = Server.CreateObject("ADODB.RecordSet")
				upd.open "Select * FROM SWEARWORDS WHERE id = "&Request.QueryString("id")&"",Connection,3,3
				upd("s_text") = s_text
				upd("s_value") = s_value
				upd.Update
			Response.Redirect "?action="
			End IF

elseIF Request.QueryString("action")= "Delete" Then
			Connection.Execute("DELETE FROM SWEARWORDS WHERE id = "&Request.QueryString("id")&"")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
elseIF Request.QueryString("action")= "ekle" Then
%>
<form action="?action=ekle2" method="post">
<table border="0" cellspacing="2" cellpadding="2" class="td_menu" align="center">
	<tr>
		<td colspan="5" align="center"><B>-=- ARGO KELÝME EKLEME ÝÞLEMLERÝ -=-</B></td>
	</tr>
	<tr>
		<td align="right">Argo Kelime : </td>
		<td><input type="text" name="s_text" class="form1"></td>
	</tr>
	<tr>
		<td align="right">Gösterilecek Deðer : </td>
		<td><input type="text" name="s_value" class="form1"></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="Ekle" class="submit"></td>
	</tr>
</table>
</form>
<%
elseIF Request.QueryString("action")= "ekle2" Then
s_text = duz(Request.Form("s_text"))
s_value = duz(Request.Form("s_value"))
			IF s_text = "" OR s_value = "" Then
			Response.Write "<div align=center class=td_menu>Tüm Alanlarý Doldurunuz !</div>"
			ELSE
			Connection.Execute("INSERT INTO SWEARWORDS (s_text,s_value) values ('"&s_text&"','"&s_value&"')")
			Response.Redirect "?action="
			End IF
End If
End IF
%>