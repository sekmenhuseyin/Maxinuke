<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Video Gonder Modul Uygulamasi											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
ScriptYolu = ""&sett("site_adres")&"/moduller.asp?mi="&Request.QueryString("mi")&""
IF Request.QueryString("eylem") = "" Then
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.gonderen.value == "") {alert("Lütfen Nickinizi Yazýnýz... !"); return false; }
if (form.buyuk.value == "") {alert("Lütfen Videonýn baþlýðýný yazýnýz... !"); return false; }
if (form.katid.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
if (form.yol.value == "") {alert("Lütfen Videonuzu yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="0" cellpadding="0" width="100%"style="<%=TableBorder%>" class=td_menu>
			<tr>
				<td colspan="2">Yotube sitesindeki hoþunuza giden beðendiðiniz video görüntülerini sitemize ekleyebilirsiniz, eklemek için yapmanýz gereken tek þey videonun adýný kategorisini ve id numarasýný yazmak, videonun id numarasý örnek olarak http://www.youtube.com/watch?v=<B>d7hzcLZ7uRc</B> adresinde izlenen bir görüntü için <B>d7hzcLZ7uRc</B> dýr yani ?v= den sonraký yazan þey.</td>
			</tr>
	<TR>
		<TD>
		<form onSubmit="return formkontrol(this)" action="?mi=<%=Request.QueryString("mi")%>&eylem=kayit_islem" method="post">
		<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
			<tr>
				<td align="right">Nickiniz : </td>
				<td><input type="text" class="form1" name="gonderen" value="<%=Session("Member")%>" size="35"></td>
			</tr>
			<tr>
				<td align="right">Videonuzun Adý :</td>
				<td><input type="text" class="form1" name="buyuk" size="35"></td>
			</tr>
			<tr>
				<td align="right">Kategorisi :</td>
				<td>
				<%set kategoricik = Connection.Execute("select * from video_cat ORDER BY isim")%>
				<select size="1" name="katid" class="form1">
				<option value="0" selected>Kategori Seçiniz</option>
				<%Do While Not kategoricik.EOF%>
				<option value="<%=kategoricik("katid")%>"><%=kategoricik("isim")%></option>
				<%kategoricik.MoveNext
				Loop%>
				</select>
				<%
				kategoricik.close
				set kategoricik = Nothing
				Connection.close
				set Connection = Nothing
				%>
				</td>
			</tr>
			<tr>
				<td align="right">Videonuzun Yolu :</td>
				<td><input type="text" class="form1" name="yol" size="35"></td>
			</tr>
			<tr>
				<td colspan="6" align="center"><input type="submit" class="submit" style="width:100%" value="Videoyu Gonder"></td>
			</tr>
		</table>
		</form>		
		</TD>
	</TR>
</TABLE>
<%
'########################################################### KAYIT ISLENIYOR ###############################################################
elseIF Request.QueryString("eylem") = "kayit_islem" Then
set rs = server.createobject("adodb.recordset")
	If Session("enter") = "Yes" then
	gonderen = Session("Member")
	else
	gonderen = request.form("gonderen")
	End if
buyuk = request.form("buyuk")
yol = request.form("yol")
katid = request.form("katid")
rs.Open "videolar", Connection, 1 , 3
rs.AddNew 
rs("buyuk") = buyuk
rs("yol") = yol
rs("onay") = False
rs("puan") = 0
rs("katilimci") = 0
rs("hit") = 0
rs("katid") = katid
rs("gonderen") = gonderen
rs.Update
rs.close
set rs = Nothing
Connection.close
set Connection = Nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=adminbilgi"
'########################################################### Patrona Haber veriliyor ###############################################################
elseIF Request.QueryString("eylem") = "adminbilgi" Then
html = html & "<div class=td_menu>Merhaba Patron<br><br>"&Now()&" tarihinde "&Request.ServerVariables("REMOTE_ADDR")&" ip adresini kullanan kisi ("&Session("Member")&") tarafýndan "&sett("site_isim")&" isimli sitenizin Videolar bölümüne yeni Video eklendi.<BR><BR>Eklenmiþ Video yayýnlanmak için onaylamanýzý beklemekte.<BR><BR>Bilgilerinize.</div>"
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" Video Eklendi"
Avanos.Body=HTML
Avanos.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
Avanos.To=""&sett("site_mail")&""
Avanos.Importance = 2
Avanos.Send
set HTML = nothing
set Avanos=nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=tesekkur"
'########################################################### TESEKKUR EDILIYOR ###############################################################
elseIF Request.QueryString("eylem") = "tesekkur" Then
%>
<meta http-equiv="refresh" content="12;url=default.asp">
<p align="center"><B>TEÞEKKÜRLER</B></p>
Göndermiþ olduðunuz Videonuz elimize ulaþtý. Yöneticiler tarafýndan onaylandýktan sonra en gec <B>3</B> gün içinde siteye eklenecektir. Eðer gitmek istediðiniz baþka bir sayfa yok ise bu sayfa 5 saniye içerisinde ana sayfaya yönlenecektir.
<%end if%>