<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												Fikra Modul Uygulamasi												#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
ScriptYolu = ""&sett("site_adres")&"/moduller.asp?mi="&Request.QueryString("mi")&""
IF Request.QueryString("eylem") = "" Then
%>
<script language="JavaScript1.2" defer>
editor_generate('metin');
</script>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.gonderen.value == "") {alert("L�tfen Nickinizi Yaz�n�z... !"); return false; }
if (form.baslik.value == "") {alert("L�tfen F�kran�n ba�l���n� yaz�n�z... !"); return false; }
if (form.cat_id.value == "0") {alert("L�tfen Kategori Se�iniz... !"); return false; }
if (form.metin.value == "") {alert("L�tfen f�kran�z� yaz�n�z... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?mi=<%=Request.QueryString("mi")%>&eylem=kayit_islem" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" width="100%" class="td_menu" style="<%=TableBorder%>">
	<tr>
		<td align="right">Nickiniz : </td>
		<td><input type="text" class="form1" name="gonderen" value="<%=Session("Member")%>" size="25"></td>
		<td align="right">F�kran�z�n Ad� :</td>
		<td><input type="text" class="form1" name="baslik" size="25"></td>
		<td align="right">Kategorisi :</td>
		<td>
		<%set kategoricik = Connection.Execute("select * from fikra_cat ORDER BY cat_name")%>
		<select size="1" name="cat_id" class="form1">
		<option value="0" selected>Kategori Se�iniz</option>
		<%Do While Not kategoricik.EOF%>
		<option value="<%=kategoricik("cat_id")%>"><%=kategoricik("cat_name")%></option>
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
		<td colspan="6" align="center"><textarea class="form1" name="metin" rows="15" style="width:100%"></textarea></td>
    </tr>
	<tr>
		<td colspan="6" align="center"><input type="submit" class="submit" style="width:100%" value="F�kray� Gonder"></td>
	</tr>
</table>
</form>
<%
'########################################################### KAYIT ISLENIYOR ###############################################################
elseIF Request.QueryString("eylem") = "kayit_islem" Then
set rs = server.createobject("adodb.recordset")
	If Session("enter") = "Yes" then
	gonderen = Session("Member")
	else
	gonderen = request.form("gonderen")
	End if
baslik = request.form("baslik")
metin = request.form("metin")
cat = request.form("cat_id")
rs.Open "fikra", Connection, 1 , 3
rs.AddNew 
rs("baslik") = baslik
rs("metin") = metin
rs("onay") = False
rs("puan") = 0
rs("katilimci") = 0
rs("hit") = 0
rs("cat") = cat
rs("gonderen") = gonderen
rs("eklenme") = Date()
rs.Update
rs.close
set rs = Nothing
Connection.close
set Connection = Nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=adminbilgi"
'########################################################### Patrona Haber veriliyor ###############################################################
elseIF Request.QueryString("eylem") = "adminbilgi" Then
html = html & "<div class=td_menu>Merhaba Patron<br><br>"&Now()&" tarihinde "&Request.ServerVariables("REMOTE_ADDR")&" ip adresini kullanan kisi ("&Session("Member")&") taraf�ndan "&sett("site_isim")&" isimli sitenizin f�kralar b�l�m�ne yeni f�kra eklendi.<BR><BR>Eklenmi� f�kra yay�nlanmak i�in onaylaman�z� beklemekte.<BR><BR>Bilgilerinize.</div>"
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" F�kra Eklendi"
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
G�ndermi� oldu�unuz F�kran�z elimize ula�t�. Y�neticiler taraf�ndan onayland�ktan sonra en gec <B>3</B> g�n i�inde siteye eklenecektir. E�er gitmek istedi�iniz ba�ka bir sayfa yok ise bu sayfa 5 saniye i�erisinde ana sayfaya y�nlenecektir.
<%end if%>