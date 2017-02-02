<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Ziyaretçi Defteri Modul Uygulamasi V 2.1									#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%IF Request.QueryString("eylem") = "" Then%>
<!--#include file="../../inc/adsense_3.asp" -->
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu">
<TR>
	<TD align="center"><A HREF="?mi=<%=Request.QueryString("mi")%>&eylem=oku"><IMG SRC="moduller/ziyaretci_defteri/oku.gif" BORDER="0" ALT="Görüþleri Oku"></A></TD>
	<TD align="center"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=yeni"><IMG SRC="moduller/ziyaretci_defteri/yaz.gif" BORDER="0" ALT="Görüþ Yaz"></a></TD>
</TR>
<TR>
	<TD align="center"><A HREF="?mi=<%=Request.QueryString("mi")%>&eylem=oku">Görüþleri Oku</A></TD>
	<TD align="center"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=yeni">Görüþ Yaz</a></TD>
</TR>
<TR>
	<TD align="center"><!--#include file="../../inc/adsense_2.asp" --></TD>
	<TD align="center"><!--#include file="../../inc/adsense_2.asp" --></TD>
</TR>
</TABLE>
<%
elseIF Request.QueryString("eylem") = "yeni" Then
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.name.value == "") {alert("Lütfen Ad Soyad Alanýný Doldurunuz..."); return false; }
if (form.email.value == "") {alert("Lütfen Email Alanýný Doldurunuz..."); return false; }
if (form.YAZI.value == "") {alert("Lütfen Görüþünüz Alanýný Doldurunuz..."); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="0" cellpadding="2" align="center"  class="td_menu">
<form name="zd" onSubmit="return formkontrol(this)" action="?mi=<%=Request.QueryString("mi")%>&eylem=Reg" method="post">
	<tr>
		<td>Ad ve Soyadýnýz : <%IF Session("Enter") = "Yes" Then%><%=Session("Member")%><%ELSE%><input type="text" name="name" size="30" class="form1"><% End IF %></td>
		<td>E-Mail Adresiniz : <%IF Session("Enter") = "Yes" Then%><%=Session("Email")%><%ELSE%><input type="text" name="email" size="30" class="form1"><%End IF%></td>
	</tr>
	<tr>
		<td colspan=4 valign="top">Görüþünüz<br><textarea name="YAZI" style="width=100%" rows="15" class="form1"></textarea></td>
	</tr>
	<tr>
		<td colspan="4" align=center><input type="submit" class="submit" value="Görüþümü Kaydet"></td>
	</tr>
</form>
</table>
<%
ElseIF Request.QueryString("eylem") = "Reg" Then
	IF Session("Enter") = "Yes" Then
	name = Session("Member")
	email = Session("Email")
	else
	name = duz(Request.Form("name"))
	email = duz(Request.Form("email"))
	End if
YAZI = duz(Request.Form("YAZI"))
SET gbEnt = Server.CreateObject("ADODB.RecordSet")
gbSQL = "Select * FROM GUESTBOOK"
gbEnt.open gbSQL,Connection,1,3
gbEnt.AddNew
gbEnt("NAME") = name
gbEnt("EMAIL") = email
gbEnt("tarih") =  Now()
gbEnt("YAZI") =  YAZI
gbEnt.Update
gbEnt.close
set gbEnt=nothing
Connection.close
set Connection=nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=haberdaret"
elseIf Request.QueryString("eylem") = "haberdaret" Then
html = html & "<div class=td_menu>Merhaba Patron<br><br>"&Now()&" tarihinde "&Request.ServerVariables("REMOTE_ADDR")&" ip adresini kullanan "&Session("Member")&" isimli üyeniz tarafýndan "&sett("site_isim")&" isimli sitenizin ziyaretçi defterine görüþ eklendi.<BR><BR>Eklenmiþ görüþ yayýnlanmak için onaylamanýzý beklemekte.<BR><BR>Bilgilerinize.</div>"
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" Ziyaretçi Görüþü Eklendi"
Avanos.Body=HTML
Avanos.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
Avanos.To=""&sett("site_mail")&""
Avanos.Importance = 2
Avanos.Send
set HTML = nothing
set Avanos=nothing
response.redirect "?mi="&Request.QueryString("mi")&"&eylem=ok"
ElseIF Request.QueryString("eylem") = "ok" Then
%>
<BR><BR><CENTER>
Sayýn <%=Session("Member")%>;<BR><BR>Görüþünüz kaydedildi, Ziyaret notlarýnýz adminlerimiz tarafýndan onaylanýr onaylanmaz yayýnlanmaya baþlayacak. Zaman ayýrýp görüþlerinizi belirttiðiniz için teþekkür ederiz.</CENTER><BR><BR><BR>
<%
elseIF Request.QueryString("eylem") = "oku" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM GUESTBOOK where onay=true ORDER BY ID DESC"
rs.open SQL,Connection,1,3
IF rs.Eof OR rs.Bof Then
Response.Write "<BR><BR><b><CENTER>Hiç Görüþ Yazýlmamýþ...</CENTER></b><BR><BR><BR>"
ELSE
Do Until rs.Eof
%>
<B><%=rs("NAME")%> - <%=rs("tarih")%></B>
<BR>
<%=rs("YAZI")%>
<hr>
<%
rs.MoveNext
Loop
End If
rs.close
set rs=nothing
Connection.close
set Connection=nothing
End IF
%>