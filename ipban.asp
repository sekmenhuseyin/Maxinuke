<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Then
	IF Request.QueryString("op") = "New" Then
%>
<form method="post" action="?op=Register">
<table border="0" cellspacing="2" cellpadding="2" align=center class="td_menu">
	<tr>
		<td colspan="2" align="center"><B>-=- ÝP ADRESÝ ENGELLEME -=-</B><hr></td>
	</tr>
	<tr>
		<td>Engellenecek IP Adresi : </td>
		<td><input type="text" name="ip" class="form1" size=50></td>
	</tr>
	<tr>
		<td>Engellenme Tarihi : </td>
		<td><%=Now()%></td>
	</tr>
	<tr>
		<td align="center" colspan="2">Sebep<br><TEXTAREA name="sebep" ROWS="15" COLS="100" class="form1"></TEXTAREA></td>
	</tr>
	<tr>
		<td align="center" colspan=2><input type="submit" value="Engelle" class="submit"></td>
	</tr>
</table>
</form>
<%
	ElseIF Request.QueryString("op") = "Register" Then
	ip = duz(Request.Form("ip"))
	sebep = duz(Request.Form("sebep"))
		If sebep = "" Then
		sebep = "Belirtilmemiþ."
		End if
	banlayan = Session("Member")
		IF ip="" Then
		Response.Write "<div align=center class=td_menu><b>Lütfen IP Adresi Yazýnýz...</b><BR><BR><input class=submit type=button value=GERI&nbsp;GIT onclick=javascript:history.back();></div>"
		ElseIF IsNumeric(ip) = False Then
		Response.Write "<div align=center class=td_menu><b>IP Adresi sayýlardan oluþmalýdýr</b><BR><BR><input class=submit type=button value=GERI&nbsp;GIT onclick=javascript:history.back();></div>"
		ELSE
		Connection.Execute("INSERT INTO BANNED_IPS (b_ip,b_date,sebep,banlayan) VALUES ('"&ip&"','"&Now()&"','"&sebep&"','"&banlayan&"')")
		Response.Redirect "ipban.asp"
		End IF
ElseIF Request.QueryString("op") = "Edit" Then
Set rs = Connection.Execute("Select * FROM BANNED_IPS WHERE b_id = "&Request.QueryString("id")&"")
ipisi=rs("b_ip")
%>
<form method="post" action="?op=Update&id=<%=Request.QueryString("id")%>">
<table border="0" cellspacing="2" cellpadding="2" align=center class="td_menu">
	<tr>
		<td colspan="2" align="center"><B>-=- ENGELLÝ DÜZENLEME -=-</B></td>
	</tr>
	<tr>
		<td>
		<%
		IF Request.QueryString("Type") = "IP" Then
		Response.Write "IP Adresi : "
		ELSE
		Response.Write "Üye Seçin : "
		End IF
		%> : 
		</td>
		<td align="left">
		<% IF Request.QueryString("Type") = "IP" Then %>
		<input type="text" name="ip" class="form1"  value="<%=ipisi%>">
		<%
		ELSE
		Set mems = Connection.Execute("Select * FROM MEMBERS ORDER BY kul_adi ASC")
		%>
		<select name="ip" size="1" class="form1">
		<%
		Do Until mems.EoF
		IF ipisi = ""&mems("uye_id")&"" Then
		opt = "selected"
		ELSE
		opt = ""
		End IF
		%>
		<option value="<%=mems("uye_id")%>" <%=opt%>><%=mems("kul_adi")%></option>
		<%
		mems.MoveNext
		Loop
		%>
		</select>
		<% End IF %>
		</td>
	</tr>
	<tr>
		<td>Tarih : </td>
		<td align="left"><%=rs("b_date")%></td>
	</tr>
	<tr>
		<td>Banlayan : </td>
		<td align="left"><%=rs("banlayan")%></td>
	</tr>
	<tr>
		<td align="center" colspan="2">Sebep<br><TEXTAREA name="sebep" ROWS="15" COLS="100" class="form1"><%=rs("sebep")%></TEXTAREA></td>
	</tr>
	<tr>
		<td align="center" colspan=2><input type="submit" value="Tamam" class="submit"></td>
	</tr>
</table>
</form>
<%
	ElseIF Request.QueryString("op") = "Update" Then
	ip = duz(Request.Form("ip"))
	sebep = duz(Request.Form("sebep"))
		IF ip="" Then
		Response.Write "<div align=center class=td_menu><b>Lütfen IP Adresi Yazýnýz...</b><BR><BR><input class=submit type=button value=GERI&nbsp;GIT onclick=javascript:history.back();></div>"
		ElseIF IsNumeric(ip) = False Then
		Response.Write "<div align=center class=td_menu><b>IP Adresi sayýlardan oluþmalýdýr</b><BR><BR><input class=submit type=button value=GERI&nbsp;GIT onclick=javascript:history.back();></div>"
		ELSE
		Connection.Execute("UPDATE BANNED_IPS Set b_ip = '"&ip&"' WHERE b_id = "&Request.QueryString("id")&"")
		Connection.Execute("UPDATE BANNED_IPS Set sebep = '"&sebep&"' WHERE b_id = "&Request.QueryString("id")&"")
		Response.Redirect "ipban.asp"
		End IF
	ElseIF Request.QueryString("op") = "Delete" Then
	Connection.Execute("DELETE FROM BANNED_IPS WHERE b_id = "&Request.QueryString("id")&"")
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ElseIF Request.QueryString("op") = "NewMember" Then
	Set mems = Connection.Execute("Select * FROM MEMBERS")
%>
<form action="?op=Register" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td><b>Üye Seçin : </b></td>
<td  align="left">
<select name="ip" size="1" class="form1">
<% Do Until mems.EoF %>
<option value="<%=mems("uye_id")%>"><%=mems("kul_adi")%></option>
<%
mems.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr>
<td ><b>Banlanma Tarihi : </b></td>
<td  align="left"><%=Now()%></td>
</tr>
	<tr>
		<td align="center" colspan="2">Sebep<br><TEXTAREA name="sebep" ROWS="15" COLS="100" class="form1"></TEXTAREA></td>
	</tr>
<tr>
<td ></td>
<td ><input type="submit" value="Tamam" class="submit"></td>
</tr>
</form>
</table>
<%
	ELSE

Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM BANNED_IPS ORDER BY b_ip DESC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="6" align="center"><div class="block_title"><B>-=- BANLANMIÞLAR LÝSTESÝ -=-</B></div></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Nick Yada ÝP</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Banlama Türü</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Banlanma Tarihi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Banlayan</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Sebep</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
IF InStr(1,rs("b_ip"),".",1) Then
kimi=rs("b_ip")
text = "Ýp Adresi"
l_text = "IP"
ELSE
set hangiuye = server.createobject("adodb.recordset")
hangiuyesql = "select * from MEMBERS where uye_id = "&rs("b_ip")&""
hangiuye.open hangiuyesql, Connection, 1, 3
kimi=hangiuye("kul_adi")
hangiuye.close
set hangiuye=nothing
text = "Üye"
l_text = "Member"
End IF
%>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td><%=kimi%></td>
		<td align=center><%=text%></td>
		<td align="center"><%=rs("b_date")%></td>
		<td align="center"><%=rs("banlayan")%></td>
		<td><%=rs("Sebep")%></td>
		<td  align="center"><a href="?op=Edit&Type=<%=l_text%>&id=<%=rs("b_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("b_id")%>">Baný Kaldýr</a></td>
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