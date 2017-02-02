<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Then
	If Request.QueryString("eylem")="" then
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.ban_type.value == "0") {alert("Reklam Cesidinizi Secmediniz !"); return false; }
if (form.ban_url.value == "") {alert("Resim yada Flash Adresi yazmadiniz !"); return false; }
if (form.go_url.value == "http://www.") {alert("Gidilecek URL yazmadiniz !"); return false; }
if (form.position.value == "0") {alert("konum Secmediniz !"); return false; }
if (form.password.value == "") {alert("Sifreyi yazmadiniz !"); return false; }
if (form.limit.value == "") {alert("kontoru yazmadiniz !"); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="2" cellpadding="2" class="td_menu" align="center" width="100%">
<form onSubmit="return formkontrol(this)" action="?eylem=yarat" method="post">
	<tr>
		<td colspan="4" align="center"><B>-=- REKLAM BANNERI EKLEME ÝÞLEMLERÝ -=-</B></td>
	</tr>
	<TR>
		<td>Reklam Bannerinizin Cesidi : </td>
		<TD>
		<SELECT name="ban_type" size="1" class="form1">
		<OPTION value=0 selected>Secim yapiniz</option>
		<OPTION value="Flash">Flash Animasyon</option>
		<OPTION value="Normal">Normal Resim</option>
		</OPTION>
		</SELECT>
		</TD>
		<td>Konum : </td>
		<td>
		<select name="position" class="form1">
		<OPTION value=0 selected>Secim yapiniz</option>
		<option value="top">Üst</option>
		<option value="bottom">Alt</option>
		<option value="blok">Blok</option>
		</select>
		</td>
	</tr>
	<TR>
		<TD>Resim yada Flash Adresi : </td>
		<TD><input type="text" name="ban_url" class="form1" size="50"></td>
		<td>Þifre : </td>
		<td><input type="text" onkeyup="sayi(this);" name="password" class="form1"></td>
	</TR>
	<TR>
		<TD>Gidilecek URL : </td>
		<TD><input type="text" name="go_url" class="form1" value="http://www." size="50"></td>
		<td>Kontör : </td>
		<td><input type="text" name="limit" class="form1" ></td>
	</tr>
	<tr>
		<td colspan="4" align="center"><input type="submit" value="Kaydet" class="submit" style="width:100%"></td>
	</tr>
</form>
</table>
<%
	elseIf Request.QueryString("eylem")="yarat" then
	ban_url = duz(Request.Form("ban_url"))
	go_url = duz(Request.Form("go_url"))
	ban_type = duz(Request.Form("ban_type"))
	position = duz(Request.Form("position"))
	password = duz(Request.Form("password"))
	limit = duz(Request.Form("limit"))
    Set kaydet = Server.CreateObject("ADODB.Recordset")
    sql = "SELECT * from BANNERS"
    kaydet.open sql, Connection, 1, 3
    kaydet.addnew
    kaydet("ban_url") = ban_url
    kaydet("go_url") = go_url
    kaydet("hit") = "0"
    kaydet("active") = true
    kaydet("position") = position
    kaydet("limit") = limit
    kaydet("password") = password
	kaydet("ban_type") = ban_type
	kaydet.update
    kaydet.close
	Response.Redirect "?eylem=liste"
	elseIf Request.QueryString("eylem")="liste" then
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open "Select * FROM BANNERS",Connection,3,3
%>
<script language="JavaScript">
<!-- // silim islemi onay uyarisi
function submitConfirm18(listForm2)
{ 
   listForm2.target="_self"; 
   listForm2.action="";
   var answer = confirm ("Silmek Istediðinize Emin misiniz ?") 
   if (answer)
       return true;
   else
       return false;
} 
//-->
</script>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="9" align="center"><B>-=- REKLAM ÝÞLEMLERÝ -=-</B></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">ID</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Banner URL</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Gidilecek URL</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tip</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Konum</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Gösterim</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Limit</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Týklanma</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
IF rs("active") = True Then
b_opt1 = "False"
b_opt2 = "Pasifleþtir"
ElseIF rs("active") = False Then
b_opt1 = "True"
b_opt2 = "<font color=red><b>Aktifleþtir</b></font>"
End IF
IF rs("position") = "top" Then
b_pos = "Üst"
ElseIF rs("position") = "bottom" Then
b_pos = "Alt"
ElseIF rs("position") = "blok" Then
b_pos = "Blok"
End IF
IF rs("ban_type") = "Flash" Then
b_type = "Flash"
ELSEIF rs("ban_type") = "Normal" Then
b_type = "Normal"
End IF
%>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td align="center"><%=rs("banner_id")%></td>
		<td><A href="?eylem=edit&id=<%=rs("banner_id")%>" title="<%=rs("ban_url")%>"><%=Left(rs("ban_url"),40)%></A></td>
		<td><A href="?eylem=edit&id=<%=rs("banner_id")%>" title="<%=rs("go_url")%>"><%=Left(rs("go_url"),40)%></a></td>
		<td align="center"><%=b_type%></td>
		<td align="center"><%=b_pos%></td>
		<td align="center"><%=rs("show")%></td>
		<td align="center"><%=rs("limit")%></td>
		<td align="center"><%=rs("hit")%></td>
		<td align="center"><a href="?eylem=Change&option=<%=b_opt1%>&id=<%=rs("banner_id")%>"><%=b_opt2%></a> | <a href="?eylem=edit&id=<%=rs("banner_id")%>">Düzenle</a> | <a onClick="return submitConfirm18(this)" href="?eylem=sil&id=<%=rs("banner_id")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
	elseIf Request.QueryString("eylem")="edit" then
	Set rs = Connection.Execute("Select * FROM BANNERS WHERE banner_id = "&Request.QueryString("id")&"")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.ban_url.value == "") {alert("Resim yada Flash Adresi yazmadiniz !"); return false; }
if (form.go_url.value == "http://www.") {alert("Gidilecek URL yazmadiniz !"); return false; }
if (form.position.value == "0") {alert("konum Secmediniz !"); return false; }
if (form.password.value == "") {alert("Sifreyi yazmadiniz !"); return false; }
if (form.limit.value == "") {alert("kontoru yazmadiniz !"); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="2" cellpadding="2" class="td_menu" align="center" width="100%">
<form action="?eylem=guncelle&id=<%=Request.QueryString("id")%>" method="post">
	<tr>
		<td colspan="4" align="center"><B>-=- REKLAM BANNERI GUNCELLEME ÝÞLEMLERÝ -=-</B></td>
	</tr>
<%
	IF rs("ban_type") = "Flash" Then
	opt1x = "selected"
	ELSEIF rs("ban_type") = "Normal" Then
	opt2x = "selected"
	End IF
	IF rs("position") = "top" Then
	opt1y = "selected"
	ElseIF rs("position") = "bottom" Then
	opt2y = "selected"
	ElseIF rs("position") = "blok" Then
	opt3y = "selected"
	End IF
%>
	<TR>
		<td>Reklam Bannerinizin Cesidi : </td>
		<TD>
		<SELECT name="ban_type" size="1" class="form1">
		<OPTION value="Flash" <%=opt1x%>>Flash Animasyon</option>
		<OPTION value="Normal" <%=opt2x%>>Normal Resim</option>
		</OPTION>
		</SELECT>
		</TD>
		<td>Konum : </td>
		<td>
		<select name="position" class="form1">
		<option value="top" <%=opt1y%>>Üst</option>
		<option value="bottom" <%=opt2y%>>Alt</option>
		<option value="blok" <%=opt3y%>>Blok</option>
		</select>
		</td>
	</tr>
	<TR>
		<TD>Resim yada Flash Adresi : </td>
		<TD><input type="text" name="ban_url" class="form1" size="50" value="<%=rs("ban_url")%>"></td>
		<td>Þifre : </td>
		<td><input type="text" onkeyup="sayi(this);" name="password" class="form1" value="<%=rs("password")%>"></td>
	</TR>
	<TR>
		<TD>Gidilecek URL : </td>
		<TD><input type="text" name="go_url" class="form1" value="<%=rs("go_url")%>" size="50"></td>
		<td>Kontör : </td>
		<td><input type="text" name="limit" class="form1" value="<%=rs("limit")%>"></td>
	</tr>
	<tr>
		<td colspan="4" align="center"><input type="submit" value="GÜNCELLE" class="submit" style="width:100%"></td>
	</tr>
</form>
</table>
<%
	elseIf Request.QueryString("eylem")="guncelle" then
ban_type = duz(Request.Form("ban_type"))
position = duz(Request.Form("position"))
ban_url = duz(Request.Form("ban_url"))
password = duz(Request.Form("password"))
go_url = duz(Request.Form("go_url"))
limit = duz(Request.Form("limit"))
SET upd = Server.CreateObject("ADODB.RecordSet")
upd.open "Select * FROM BANNERS WHERE banner_id = "&Request.QueryString("id")&"",Connection,3,3
upd("ban_type") = ban_type
upd("position") = position
upd("ban_url") = ban_url
upd("go_url") = go_url
upd("password") = password
upd("limit") = limit
upd.Update
Response.Redirect "?eylem=liste"
	elseIf Request.QueryString("eylem")="Change" then
Connection.Execute("UPDATE BANNERS SET active = "&Request.QueryString("option")&" WHERE banner_id = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	elseIf Request.QueryString("eylem")="sil" Then
Connection.Execute("DELETE FROM BANNERS WHERE banner_id = "&Request.QueryString("id")&"")
Response.Redirect "?eylem=liste"
	End if
End IF
%>