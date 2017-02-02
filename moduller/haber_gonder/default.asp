<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Haber Ekle Modul Uygulamasi V 3.0										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<!--#include file="../../fckeditor/fckeditor.asp" -->
<%
If Request.QueryString("eylem") = "" Then
Set cats = Connection.Execute("Select * FROM NEWS_CATS ORDER BY cat_name ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function giris_kontrol(form) 
{
if (form.ntitle.value == "") {alert("Lütfen Haber Baþlýðýný Yazýnýz"); return false; }
if (form.ncat.value == "0") {alert("Lütfen Kategori Seçiniz"); return false; }
if (form.ntext.value == "") {alert("Lütfen Haber Metnini Yazýnýz"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return giris_kontrol(this)" action="?mi=<%=Request.QueryString("mi")%>&eylem=haberekle" method="post">
<table border="0" cellspacing="3" cellpadding="3" width="100%" align="center" class="td_menu">
	<tr>
		<td>Baþlýk : <input type="text" name="ntitle" size="50" class="form1" maxlength="50"></td>
		<td>Kategori : 
		<select name="ncat" size="1" class="form1">
		<option value="0" selected>Lutfen Seciniz</option>
		<%
		Do Until cats.Eof
		Response.Write "<option value="""&cats("cat_id")&""">"&cats("cat_name")&"</option>"
		cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="2">
		<%
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.BasePath = "FCKeditor/"
		oFCKeditor.Value	= ""
		oFCKeditor.Create "ntext"
		%>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" style="width=100%" value="Haberimi Gönder" class="submit"></td>
	</tr>
</table>
</form>
<%
cats.Close
Set cats = Nothing
Connection.Close
Set Connection = Nothing
elseIf Request.QueryString("eylem") = "haberekle" Then
title = duz(Request.Form("ntitle"))
news = duz(Request.Form("ntext"))
cat = Request.Form("ncat")

If title = "" Or  news = "" Or  cat = "0" then
Response.Write WriteMsg("Haber Baþlýðý, Kategorisi yada Metni eksik kaldi")
else
	If Session("Level") = "1" Or Session("Level") = "4" then
	onay = true
	else
	onay = False
	End if
	If session("gk")="" then
	gk="000000000"
	else
	gk=session("gk")
	End if
	If session("Member")="" then
	Member="Anonim"
	else
	Member=session("Member")
	End if
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM NEWS"
ent.open entSQL,Connection,1,3
ent.AddNew
ent("baslik") = title
ent("metin") = news
ent("ekleyen") = Member
ent("tarih") = Now()
ent("onay") = onay
ent("puan") = 0
ent("katilimci") = 0
ent("okunma") = 0
ent("cat") = cat
ent("gk") = gk
ent.Update
ent.Close
Set ent = Nothing
Connection.Close
Set Connection = Nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=haberdaret"
End if
elseIf Request.QueryString("eylem") = "haberdaret" Then
html = html & "<div class=td_menu>Merhaba Patron<br><br>"&Now()&" tarihinde "&Request.ServerVariables("REMOTE_ADDR")&" ip adresini kullanan biri tarafýndan "&sett("site_isim")&" isimli sitenize haber eklendi.<BR><BR>Eklenmiþ haber onaysýzlar kutusunda bekletilmekte.<BR><BR>Bilgilerinize.</div>"
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" Haber Eklendi"
Avanos.Body=HTML
Avanos.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
Avanos.To=""&sett("site_mail")&""
Avanos.Importance = 2
Avanos.Send
set HTML = nothing
set Avanos=nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=ok"
elseIf Request.QueryString("eylem") = "ok" Then
	If Session("Level") = "1" Or Session("Level") = "4" then
	Response.Write "Sayin Yönetici<BR><BR>Göndermiþ olduðunuz haber yayýnlanmaya baþladý.<BR><BR>Bilgileribinize."
	else
	Response.Write "Göndermiþ olduðunuz haber elimize ulaþtý, admin yada haber editörlerimiz tarafýndan incelenip en yakýn zamanda sitedeki yerini alacaktýr.<BR><BR>Ýlginizden ötürü teþekkür ederiz..."
	End if
End IF
%>