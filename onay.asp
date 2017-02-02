<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA
action = Request.QueryString("action")
action = QS_CLEAR(action)
if action = "" then
call onaysizsin
elseif action = "onayla" then
call onay1
elseif action = "fel" then
call hata
elseif action = "ok" then
call tamam
elseif action = "tekrar" then
call birdaha
elseif action = "tekrarr" then
call birdahaa
elseif action = "okk" then
call tamamm
end if

Sub onaysizsin
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" bgcolor="<%=content_bg%>" class="td_menu">
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" align="center"><B>ONAYSIZ ÜYE</B></font></td>
	</tr>
	<tr>
		<td valign="top">Sayýn üye adayýmýz; sitemizin tüm alanlarýndan faydalanabilmeniz için üyeliðinizi aktifleþtirmeniz gerekmektedir, siz henüz üyeliðinizi onaylatmamýþsýnýz. Lütfen kayýt esnasýnda kullanmýþ olduðunuz mail adresini kontrol ediniz. Bazý durumlarda Hotmail, Yahoo, MSN, Mynet vb. mail saðlayýcýlarýnda aktive maili reklam kutusuna (Junk Mail) gidebilmektedir, lütfen bu alanlarýda kontrol ediniz. Eðer aktive mailini almadýysanýz alt kýsýmdan AKTIVE MAILIMI YENILE butonunu kullanýnýz, yada mail adresiniz doðru ama bizden kaynaklanan bir sorun olduðundan dolayý aktive olamadýðýnýzý düþünüyorsanýz <A HREF="mailto:<%=sett("site_mail")%>?subject=Mail Aktivite Hatasý"><%=sett("site_mail")%></A> adresinden bizimle iletiþim kurabilirsiniz. Eger yeni bir mail adresi ile hesap açmak istiyorsanýz YENÝ KAYIT butonunu kullanýnýz.<BR><BR>Saygýlarýmýzla <%=sett("site_isim")%> Ekibi.</td>
	</tr>
	<tr>
		<td valign="top" align="center" ><a href="?action=tekrar"><button onclick="parentElement.click();" class=submit>AKTIVE MAILIMI YENILE</button></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="uye_islem.asp?action=new"><button onclick="parentElement.click();" class=submit>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;YENÝ KAYIT&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</button></a><BR><BR></td>
	</tr>
</table>
<%
End Sub

Sub onay1
gk = Request.QueryString("gk")
If gk = "" OR IsNumeric(gk) = False Then
response.redirect "?action=fel"
response.End
End if
Set onayver=Server.CreateObject("Adodb.Recordset")
onayver.ActiveConnection = Connection
Sorgu="SELECT * FROM MEMBERS WHERE gk="& gk &""
onayver.open Sorgu,Connection,1,3
if onayver.eof then
response.redirect "?action=fel"
response.end
end if
onayver.update "onay","True"
onayver.close
set onayver=nothing
Connection.close
set Connection=nothing
Response.Redirect("?action=ok")
End Sub

Sub tamam
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr>
		<td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" align="center"><B>AKTIVE TAMAM</B></font></td>
	</tr>
	<tr>
		<td valign="top" bgcolor="<%=content_bg%>" class="td_menu">&nbsp;<BR>Üyeliðiniz baþarýyla onaylandý, artýk üst kýsýmý kullanarak sisteme giriþ yapabilir ve tüm alanlarý kullanabilirsiniz.<BR><BR></td>
	</tr>
</table>
<%
End Sub

Sub hata
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr>
		<td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" align="center"><B>HATA</B></font></td>
	</tr>
	<tr>
		<td valign="top" bgcolor="<%=content_bg%>" class="td_menu">Üzgünüz güvenlik linkiniz doðru yada tutarlý görünmüyor, lütfen size göndermiþ olduðumuz maili inceliyerek tekrar deneyin.</td>
	</tr>
</table>
<%
End sub

Sub birdaha
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
	<tr>
		<td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>AKTÝVÝTE YENÝLEME</B></CENTER></td>
	</tr>
	<tr>
		<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu"><BR><BR>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.mail.value == "") {alert("Lutfen kayýt esnasýnda belýrttýgýnýz maýl adresýnýzý yazýnýz."); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=tekrarr" method="post">
Mail Adresiniz<BR><BR><input type="text" name="mail" size="30" class="form1" maxlength="40"><BR><BR><input type="submit" value="<%=lost_continue%>" class="submit"></form><BR><BR>
		</TD>
	</TR>
</TABLE>
<%
End sub

Sub birdahaa
email = Request.Form("mail")
email = QS_CLEAR(email)
Set uyectrl = Connection.Execute("Select * FROM MEMBERS WHERE onay=false and email = '"&email&"'")
IF uyectrl.eof Then
Response.Write WriteMsg("Girdiðiniz Mail adresi sistemde onay bekleyen kisiler arasinda bulunamadý.")
ELSE
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.To = uyectrl("email")
objCDO.From = ""&sett("site_isim")&"<"&sett("site_mail")&">"
objCDO.Subject = ""&sett("site_isim")&" Uyelik Aktivitasyonu"
objCDO.BodyFormat=0
objCDO.MailFormat=0
govde=govde &" <font face=""Tahoma"" style=""font-size:11px;""> "& chr(10)
govde=govde &" Sayin <b>"&uyectrl("isim")&"</b>; "& chr(10)
govde=govde &" <br> "& chr(10)
govde=govde &" <b>"&sett("site_isim")&"</b> isimli sitedeki üyeliðinizi aktifleþtirmek için yeniden talepte bulundunuz lütfen alt kýsýmdaki linki týlayýn."& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &"<a href="""&sett("site_adres")&"/onay.asp?action=onayla&gk="&uyectrl("gk")&""">UYELIGIMI ONAYLA</a>"& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" <b><a href="""&sett("site_adres")&""">"&sett("site_isim")&"</a></b> (Powered by <a href=""http://www.maxinuke.com"">Maxi Nuke</a>) "& chr(10)
govde=govde &"</font>"& chr(10)
objCDO.Body = govde
objCDO.Importance = 2
objCDO.Send
Set objCDO = Nothing
response.redirect "?action=okk"
End If
uyectrl.Close
Set uyectrl = Nothing
end sub

Sub tamamm
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
	<tr>
		<td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" align="center"><B>AKTIVE YENILENDI</B></font></td>
	</tr>
	<tr>
		<td valign="top" bgcolor="<%=content_bg%>" class="td_menu">Sayýn üye adayýmýz; üyeliðinizi aktifleþtirmeniz için gereken maili tekrardan yolladýk. Lütfen mail adresini kontrol ediniz. Bazý durumlarda Hotmail, Yahoo, MSN, Mynet vb. mail saðlayýcýlarýnda aktive maili reklam kutusuna (Junk Mail) gidebilmektedir, lütfen bu alanlarýda kontrol ediniz. Saygýlarýmýzla <%=sett("site_isim")%> Ekibi.<BR><BR></td>
	</tr>
</table>
<%
End sub
call ORTA2
call ALT
%>