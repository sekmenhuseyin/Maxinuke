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
if (form.gonderen.value == "") {alert("L�tfen Nickinizi Yaz�n�z... !"); return false; }
if (form.buyuk.value == "") {alert("L�tfen Videon�n ba�l���n� yaz�n�z... !"); return false; }
if (form.katid.value == "0") {alert("L�tfen Kategori Se�iniz... !"); return false; }
if (form.yol.value == "") {alert("L�tfen Videonuzu yaz�n�z... !"); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="0" cellpadding="0" width="100%"style="<%=TableBorder%>" class=td_menu>
			<tr>
				<td colspan="2">Yotube sitesindeki ho�unuza giden be�endi�iniz video g�r�nt�lerini sitemize ekleyebilirsiniz, eklemek i�in yapman�z gereken tek �ey videonun ad�n� kategorisini ve id numaras�n� yazmak, videonun id numaras� �rnek olarak http://www.youtube.com/watch?v=<B>d7hzcLZ7uRc</B> adresinde izlenen bir g�r�nt� i�in <B>d7hzcLZ7uRc</B> d�r yani ?v= den sonrak� yazan �ey.</td>
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
				<td align="right">Videonuzun Ad� :</td>
				<td><input type="text" class="form1" name="buyuk" size="35"></td>
			</tr>
			<tr>
				<td align="right">Kategorisi :</td>
				<td>
				<%set kategoricik = Connection.Execute("select * from video_cat ORDER BY isim")%>
				<select size="1" name="katid" class="form1">
				<option value="0" selected>Kategori Se�iniz</option>
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
html = html & "<div class=td_menu>Merhaba Patron<br><br>"&Now()&" tarihinde "&Request.ServerVariables("REMOTE_ADDR")&" ip adresini kullanan kisi ("&Session("Member")&") taraf�ndan "&sett("site_isim")&" isimli sitenizin Videolar b�l�m�ne yeni Video eklendi.<BR><BR>Eklenmi� Video yay�nlanmak i�in onaylaman�z� beklemekte.<BR><BR>Bilgilerinize.</div>"
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
<p align="center"><B>TE�EKK�RLER</B></p>
G�ndermi� oldu�unuz Videonuz elimize ula�t�. Y�neticiler taraf�ndan onayland�ktan sonra en gec <B>3</B> g�n i�inde siteye eklenecektir. E�er gitmek istedi�iniz ba�ka bir sayfa yok ise bu sayfa 5 saniye i�erisinde ana sayfaya y�nlenecektir.
<%end if%>