<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Reklam Kontrolu Modul Uygulamasi V 2.0										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
IF Request.QueryString("eylem") = "Cikis" Then
Session("R_V_girdimi") = ""
Session("banner_id") = ""
Response.Redirect "?mi="&Request.QueryString("mi")&""
ElseIF Request.QueryString("eylem") = "Giris" Then
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.banner_id.value == "") {alert("Lutfen Reklam numaranizi yaziniz"); return false; }
if (form.password.value == "") {alert("Lutfen Reklamveren sifrenizi yaziniz"); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="2" cellpadding="1" width="100%" align="center" class="td_menu">
<form method="post" onSubmit="return formkontrol(this)" action="?mi=<%=Request.QueryString("mi")%>&eylem=GirisYap">
	<tr>
		<td width="40%" align="right">Reklam No : </td>
		<td width="60%"><input type="text"  onkeyup="sayi(this);" name="banner_id" size="20" class="form1"></td>
	</tr>
	<tr>
		<td width="40%" align="right">Reklamveren Þifreniz : </td>
		<td width="60%"><input type="password"  onkeyup="sayi(this);" name="password" size="20" class="form1"></td>
</tr>
<tr>
<td width="40%"></td>
<td width="60%"><input type="submit" value="Giriþ" class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF Request.QueryString("eylem") = "GirisYap" Then
banner_id = Request.Form("banner_id")
password = duz(Request.Form("password"))
set RK = Connection.Execute("SELECT * FROM BANNERS")
Do Until RK.Eof
	If ucase(banner_id) = ucase(RK("banner_id")) then
	ban_bid = True
		If ucase(password) = ucase(RK("password")) then 
		ban_pass = True
Session("banner_id") = RK("banner_id")
Session("R_V_girdimi") = "Yes"
Response.Redirect "?mi="&Request.QueryString("mi")&""
		End If
	End If
RK.MoveNext
Loop
	If ban_bid <> True then
	Response.Write "<BR><BR><BR><B><CENTER>Yazdýðýnýz Reklam No.'suna ait reklam bulunamadý</CENTER></B><BR><BR><BR>"
	%>
	<CENTER><!--#include file="../../INC/geri.asp" --></CENTER><BR><BR><BR>
	<%
	ElseIf ban_pass <> True then
	Response.Write "<BR><BR><BR><CENTER><B>Yazdýðýnýz Þifre Yanlýþ</B></CENTER><BR><BR><BR>"
	%>
	<CENTER><!--#include file="../../INC/geri.asp" --></CENTER><BR><BR><BR>
	<%
	End If
RK.Close
Set RK = Nothing
ELSEIF Session("R_V_girdimi") = "Yes" Then
Set RK = Server.CreateObject("ADODB.RecordSet")
RK.open "Select * FROM BANNERS WHERE banner_id = "&Session("banner_id")&"",Connection,3,3
%>
<table border="0" cellspacing="3" cellpadding="0" align="center" width="100%" class="td_menu" style="<%=TableBorder%>">
	<tr>
		<td colspan="4" align="center">
<%If RK("ban_type") = "Flash" then%>
		<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" WIDTH="468" HEIGHT="60"><param name="movie" value="<%=RK("ban_url")%>"></object>
<%elseIf RK("ban_type") = "Normal" then%>
<A HREF="<%=RK("go_url")%>"><img src="<%=RK("ban_url")%>" border="0" width="468"></A>
<%End if%>
		</td>
	</tr>
	<tr>
		<td align="right"><B>Reklam No</B> : </td>
		<td><%=RK("banner_id")%></td>
		<td align="right"><B>Reklam Durumu</B> : </td>
		<td><%If RK("active") = True Then response.write "Yayinlaniyor" Else response.write "Beklemede" End if%></td>
	</tr>
	<tr>
		<td align="right"><B>Alýnan Kontör</B> : </td>
		<td><%=RK("limit")%></td>
		<td align="right"><B>Harcanan Kontör</B> : </td>
		<td><%=RK("show")%></td>
	</tr>
	<tr>
		<td align="right"><B>Pozisyon</B> : </td>
		<td>
		<%
		If RK("position") = "top" Then
		response.write "Ust Alanda"
		ElseIf RK("position") = "bottom" Then
		response.write "Alt Alanda"
		ElseIf RK("position") = "yan" Then
		response.write "Yan Blokda"
		End if%></td>
		<td align="right"><B>Reklam Turu</B> : </td>
		<td><%=RK("ban_type")%></td>
	</tr>
	<tr>
		<td align="right"><B>Kalan Kontör</B> : </td>
		<td><%=Int(RK("limit")-RK("show"))%></td>
		<td align="right"><B>Týklanma</B> : </td>
		<td><%=RK("hit")%></td>
	</tr>
<TR>
	<TD colspan="4" align="center">Reklam sürenizi uzatmak yada olasý isteklerinizi paylasmak için <%=sett("site_isim")%> site yöneticilerine <A HREF="<%=sett("site_mail")%>"><%=sett("site_mail")%></A> adresine mail yollayabilrsiniz.</TD>
</TR>
</TABLE>
<%
RK.Close
Set RK = Nothing
ELSE
Response.Redirect "?mi="&Request.QueryString("mi")&"&eylem=Giris"
End IF
%>