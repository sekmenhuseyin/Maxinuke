<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Iletisim Formu Modul Uygulamasi V 1.0										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%IF Request.QueryString("eylem") = "" Then%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.ad.value == "") {alert("Adý Soyad yada Nickiniz alani bos !"); return false; }
if (form.email.value == "") {alert("E-Mail Adresiniz alani bos !"); return false; }
if (form.mesaj.value == "") {alert("Mesajiniz alani bos !"); return false; }
if (form.guvenlik.value == "") {alert("Lütfen Güvenlik Kodunu yazýnýz!"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" name="form" method="post" action="?mi=<%=Request.QueryString("mi")%>&eylem=gonder">
<%
Randomize
a1 =  Int(Rnd *9) +1
a2 =  Int(Rnd *9) +1
a3 =  Int(Rnd *9) +1
a4 =  Int(Rnd *9) +1
a5 =  Int(Rnd *9) +1
%>
<input type="hidden" name="gguvenlik" value="<%=a1%><%=a2%><%=a3%><%=a4%><%=a5%>">
<table border="0" cellpadding="3" cellspacing="3" class="td_menu" align="center">
	<tr>
		<td ALIGN="RIGHT">Adý Soyad yada Nickiniz : </td>
		<td><input name="ad" type="text" size="50" class="form1" VALUE="<%=Session("Member")%>"></td>
	</tr>
	<tr>
		<td ALIGN="RIGHT">E-Mail Adresiniz : </td>
		<td><input name="email" type="text" size="50" class="form1" VALUE="<%=Session("Email")%>"></td>
	</tr>
	<tr>
		<td ALIGN="RIGHT">Telefonunuz : </td>
		<td><input name="tel" type="text" size="50" class="form1" onkeyup="sayi(this);"></td>
	</tr>
	<tr>
		<td ALIGN="RIGHT">Mesajiniz : </td>
		<td><textarea name="mesaj" cols="50%" class="form1" rows="10"></textarea></td>
	</tr>
	<tr>
		<td ALIGN="RIGHT">Kod : </td>
		<td><%=a1%><%=a2%><%=a3%><%=a4%><%=a5%><input name="guvenlik" type="text" size="10" class="form1"></td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" name="Submit" value="Mesajimi Gonder" class="Submit"></td>
	</tr>
</table>
</form>
<%
elseIF Request.QueryString("eylem") = "gonder" Then
	If not Request.Form("gguvenlik") = Request.Form("guvenlik") Then
	response.write WriteMsg(sc_err2)
	else
	If Session("Enter") = "Yes" Then
	kisi = Session("Member")
	maili= Session("Email")
	Else
	kisi = Request.Form("ad")
	maili = Request.Form("email")
	End if
html = html & "<div class=td_menu>Merhaba Patron<br><br>"&Now()&" tarihinde "&Request.ServerVariables("REMOTE_ADDR")&" ip adresini kullanan "&kisi&" kisi tarafýndan "&sett("site_isim")&" isimli sitenizin iletisim formu kullanilarak yonetiminize mail yollandi.<hr>Ad Soyad : "&kisi&"<br>Mail Adresi : "&maili&"<br>Telefon : "&Request.Form("tel")&"<br>"&Request.Form("mesaj")&"</div>"
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" Iletisim Formu Iletisi"
Avanos.Body=HTML
Avanos.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
Avanos.To=""&sett("site_mail")&""
Avanos.Value("Geri") = "&maili&"
Avanos.Importance = 2
Avanos.Send
set HTML = nothing
set Avanos=nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=ok"
End if
ElseIF Request.QueryString("eylem") = "ok" Then
%>
<BR><BR><CENTER>
Sayýn <%=Session("Member")%>;<BR><BR>Ýletiniz site yönetimine gönderildi, en yakýn zamanda ilgili departmanýmýzdan yanýt alacaksýnýz.</CENTER><BR><BR><BR>
<%end if%>