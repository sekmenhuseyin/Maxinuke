<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Ziyaret�i Defteri Modul Uygulamasi V 2.1									#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%IF Request.QueryString("eylem") = "" Then%>
<!--#include file="../../inc/adsense_3.asp" -->
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu">
<TR>
	<TD align="center"><A HREF="?mi=<%=Request.QueryString("mi")%>&eylem=oku"><IMG SRC="moduller/ziyaretci_defteri/oku.gif" BORDER="0" ALT="G�r��leri Oku"></A></TD>
	<TD align="center"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=yeni"><IMG SRC="moduller/ziyaretci_defteri/yaz.gif" BORDER="0" ALT="G�r�� Yaz"></a></TD>
</TR>
<TR>
	<TD align="center"><A HREF="?mi=<%=Request.QueryString("mi")%>&eylem=oku">G�r��leri Oku</A></TD>
	<TD align="center"><a href="?mi=<%=Request.QueryString("mi")%>&eylem=yeni">G�r�� Yaz</a></TD>
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
if (form.name.value == "") {alert("L�tfen Ad Soyad Alan�n� Doldurunuz..."); return false; }
if (form.email.value == "") {alert("L�tfen Email Alan�n� Doldurunuz..."); return false; }
if (form.YAZI.value == "") {alert("L�tfen G�r���n�z Alan�n� Doldurunuz..."); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="0" cellpadding="2" align="center"  class="td_menu">
<form name="zd" onSubmit="return formkontrol(this)" action="?mi=<%=Request.QueryString("mi")%>&eylem=Reg" method="post">
	<tr>
		<td>Ad ve Soyad�n�z : <%IF Session("Enter") = "Yes" Then%><%=Session("Member")%><%ELSE%><input type="text" name="name" size="30" class="form1"><% End IF %></td>
		<td>E-Mail Adresiniz : <%IF Session("Enter") = "Yes" Then%><%=Session("Email")%><%ELSE%><input type="text" name="email" size="30" class="form1"><%End IF%></td>
	</tr>
	<tr>
		<td colspan=4 valign="top">G�r���n�z<br><textarea name="YAZI" style="width=100%" rows="15" class="form1"></textarea></td>
	</tr>
	<tr>
		<td colspan="4" align=center><input type="submit" class="submit" value="G�r���m� Kaydet"></td>
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
html = html & "<div class=td_menu>Merhaba Patron<br><br>"&Now()&" tarihinde "&Request.ServerVariables("REMOTE_ADDR")&" ip adresini kullanan "&Session("Member")&" isimli �yeniz taraf�ndan "&sett("site_isim")&" isimli sitenizin ziyaret�i defterine g�r�� eklendi.<BR><BR>Eklenmi� g�r�� yay�nlanmak i�in onaylaman�z� beklemekte.<BR><BR>Bilgilerinize.</div>"
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" Ziyaret�i G�r��� Eklendi"
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
Say�n <%=Session("Member")%>;<BR><BR>G�r���n�z kaydedildi, Ziyaret notlar�n�z adminlerimiz taraf�ndan onaylan�r onaylanmaz yay�nlanmaya ba�layacak. Zaman ay�r�p g�r��lerinizi belirtti�iniz i�in te�ekk�r ederiz.</CENTER><BR><BR><BR>
<%
elseIF Request.QueryString("eylem") = "oku" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM GUESTBOOK where onay=true ORDER BY ID DESC"
rs.open SQL,Connection,1,3
IF rs.Eof OR rs.Bof Then
Response.Write "<BR><BR><b><CENTER>Hi� G�r�� Yaz�lmam��...</CENTER></b><BR><BR><BR>"
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