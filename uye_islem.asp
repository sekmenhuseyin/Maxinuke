<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA
action = Request.QueryString("action")
action = QS_CLEAR(action)
if action = "new" then
call mem_new
elseif action = "register" then
call reg
elseif action = "lostpass" then
call lostpassword
end if


Sub mem_new
reg_sec_code = CPASS(13)
Session("RegSecurityCode") = reg_sec_code
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.kuladi.value == "") {alert("<%=error1%>"); return false; }
if (form.password.value == "") {alert("<%=error2%>"); return false; }
if (form.email.value == "") {alert("<%=error3%>"); return false; }
if (form.isim.value == "") {alert("<%=error4%>"); return false; }
if (form.g_soru.value == "") {alert("<%=error5%>"); return false; }
if (form.g_cevap.value == "") {alert("<%=error6%>"); return false; }
if (form.sehir.value == "0") {alert("L�tfen Yasadiginiz Sehri Belirtiniz"); return false; }
if (form.cinsiyet.value == "0") {alert("L�tfen Cinsiyetinizi Belirtiniz"); return false; }
if (form.yas_3.value == "1940") {alert("L�tfen Ya��n�z� Belirtiniz"); return false; }
if (form.security_code.value == "") {alert("<%=sc_err1%>"); return false; }
return true;
}
</SCRIPT>

<SCRIPT type=text/javascript>
var checkobj
function agreesubmit(el){
checkobj=el
if (document.all||document.getElementById)
	{
	for (i=0;i<checkobj.form.length;i++)
		{  //hunt down submit button
		var tempobj=checkobj.form.elements[i]
		if(tempobj.type.toLowerCase()=="submit")
		tempobj.disabled=!checkobj.checked
		}
	}
}
</SCRIPT>
<script>
var degerler="1234567890"
	function koduret(plength)
		{
		sabit=''
		for (i=0;i<plength;i++)
		sabit+=degerler.charAt(Math.floor(Math.random()*degerler.length))
		return sabit
		}
	function populateform(enterlength)
		{
		document.kayit.gk.value=koduret(enterlength)
		}
</script>
<form name="kayit" onSubmit="return formkontrol(this)" action="?action=register" method="post">
<input type="hidden" name="uzunluk" value="9">
<input type="hidden" name="gk">
			  <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=m_title%></B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table width="100%" border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
  <tr bgcolor="<%=forum_bg_1%>"> 
    <td width="50%">&nbsp;<%=uye_kuladi%></td><td width="50%"><input type="text" name="kuladi" class="form1" size="30" maxlength="15"> Max: 15 Karakter</td>
  </tr>
  <tr bgcolor="<%=forum_bg_2%>">
    <td width="50%">&nbsp;<%=uye_password%></td><td width="50%"><input type="text" name="password" class="form1" size="30"></td>
  </tr>
  <tr bgcolor="<%=forum_bg_1%>"> 
    <td width="50%">&nbsp;Mail Adresi</td><td width="50%"><input type="text" name="email" class="form1" size="30"></td>
  </tr>
  <tr bgcolor="<%=forum_bg_2%>">
    <td width="50%">&nbsp;Ad ve Soyad</td><td width="50%"><input type="text" name="isim" class="form1" size="30"></td>
  </tr>
  <tr bgcolor="<%=forum_bg_1%>">
    <td width="50%">&nbsp;Gizli Soru</td><td width="50%"><input type="text" name="g_soru" class="form1" size="30"></td>
  </tr>
  <tr bgcolor="<%=forum_bg_2%>">
    <td width="50%">&nbsp;<%=register_answer%></td><td width="50%"><input type="text" name="g_cevap" class="form1" size="30"></td>
  </tr>
  <tr bgcolor="<%=forum_bg_2%>">
    <td width="50%">&nbsp;�ehir</td><td width="50%">
	<SELECT name="sehir" class="form1" >
	<OPTION value=0 selected>- L�tfen �ehir Se�imi Yap�n�z. -</OPTION>
	<OPTION value="Adana">Adana</OPTION>
	<OPTION value="Ad�yaman">Ad�yaman</OPTION>
	<OPTION value="Afyon">Afyon</OPTION>
	<OPTION value="A�r�">A�r�</OPTION>
	<OPTION value="Aksaray">Aksaray</OPTION>
	<OPTION value="Amasya">Amasya</OPTION>
	<OPTION value="Ankara">Ankara</OPTION>
	<OPTION value="Antalya">Antalya</OPTION>
	<OPTION value="Ardahan">Ardahan</OPTION>
	<OPTION value="Artvin">Artvin</OPTION>
	<OPTION value="Ayd�n">Ayd�n</OPTION>
	<OPTION value="Bal�kesir">Bal�kesir</OPTION>
	<OPTION value="Bart�n">Bart�n</OPTION>
	<OPTION value="Batman">Batman</OPTION>
	<OPTION value="Bayburt">Bayburt</OPTION>
	<OPTION value="Bilecik">Bilecik</OPTION>
	<OPTION value="Bing�l">Bing�l</OPTION>
	<OPTION value="Bitlis">Bitlis</OPTION>
	<OPTION value="Bolu">Bolu</OPTION>
	<OPTION value="Burdur">Burdur</OPTION>
	<OPTION value="Bursa">Bursa</OPTION>
	<OPTION value="�anakkale">�anakkale</OPTION>
	<OPTION value="�ank�r�">�ank�r�</OPTION>
	<OPTION value="�orum">�orum</OPTION>
	<OPTION value="Denizli">Denizli</OPTION>
	<OPTION value="Diyarbak�r">Diyarbak�r</OPTION>
	<OPTION value="D�zce">D�zce</OPTION>
	<OPTION value="Edirne">Edirne</OPTION>
	<OPTION value="Elaz��">Elaz��</OPTION>
	<OPTION value="Erzincan">Erzincan</OPTION>
	<OPTION value="Erzurum">Erzurum</OPTION>
	<OPTION value="Eski�ehir">Eski�ehir</OPTION>
	<OPTION value="Gaziantep">Gaziantep</OPTION>
	<OPTION value="Gazimagosa">Gazimagosa</OPTION>
	<OPTION value="Giresun">Giresun</OPTION>
	<OPTION value="Girne">Girne</OPTION>
	<OPTION value="G�m��hane">G�m��hane</OPTION>
	<OPTION value="G�zelyurt">G�zelyurt</OPTION>
	<OPTION value="Hakkari">Hakkari</OPTION>
	<OPTION value="Hatay">Hatay</OPTION>
	<OPTION value="I�d�r">I�d�r</OPTION>
	<OPTION value="Isparta">Isparta</OPTION>
	<OPTION value="Mersin">��el (Mersin)</OPTION>
	<OPTION value="�skele">�skele</OPTION>
	<OPTION value="�stanbul-Anadolu">�stanbul-Anadolu</OPTION>
	<OPTION value="�stanbul-Avrupa">�stanbul-Avrupa</OPTION>
	<OPTION value="�zmir">�zmir</OPTION>
	<OPTION value="Kahramanmara�">Kahramanmara�</OPTION>
	<OPTION value="Karab�k">Karab�k</OPTION>
	<OPTION value="Karaman">Karaman</OPTION>
	<OPTION value="Kars">Kars</OPTION>
	<OPTION value="Kastamonu">Kastamonu</OPTION>
	<OPTION value="Kayseri">Kayseri</OPTION>
	<OPTION value="K�r�kkale">K�r�kkale</OPTION>
	<OPTION value="K�rklareli">K�rklareli</OPTION>
	<OPTION value="K�r�ehir">K�r�ehir</OPTION>
	<OPTION value="Kilis">Kilis</OPTION>
	<OPTION value="Kocaeli">Kocaeli</OPTION>
	<OPTION value="Konya">Konya</OPTION>
	<OPTION value="K�tahya">K�tahya</OPTION>
	<OPTION value="Lefkosa">Lefkosa</OPTION>
	<OPTION value="Malatya">Malatya</OPTION>
	<OPTION value="Manisa">Manisa</OPTION>
	<OPTION value="Mardin">Mardin</OPTION>
	<OPTION value="Mu�la">Mu�la</OPTION>
	<OPTION value="Mu�">Mu�</OPTION>
	<OPTION value="Nev�ehir">Nev�ehir</OPTION>
	<OPTION value="Ni�de">Ni�de</OPTION>
	<OPTION value="Ordu">Ordu</OPTION>
	<OPTION value="Osmaniye">Osmaniye</OPTION>
	<OPTION value="Rize">Rize</OPTION>
	<OPTION value="Sakarya">Sakarya</OPTION>
	<OPTION value="Samsun">Samsun</OPTION>
	<OPTION value="Siirt">Siirt</OPTION>
	<OPTION value="Sinop">Sinop</OPTION>
	<OPTION value="Sivas">Sivas</OPTION>
	<OPTION value="�anl�urfa">�anl�urfa</OPTION>
	<OPTION value="��rnak">��rnak</OPTION>
	<OPTION value="Tekirda�">Tekirda�</OPTION>
	<OPTION value="Tokat">Tokat</OPTION>
	<OPTION value="Trabzon">Trabzon</OPTION>
	<OPTION value="Tunceli">Tunceli</OPTION>
	<OPTION value="U�ak">U�ak</OPTION>
	<OPTION value="Van">Van</OPTION>
	<OPTION value="Yalova">Yalova</OPTION>
	<OPTION value="Yozgat">Yozgat</OPTION>
	<OPTION value="Zonguldak">Zonguldak</OPTION>
	<OPTION value="Yurtdisi">Yurtdisi</OPTION>
	</SELECT>
	</td>
  </tr>
  <tr bgcolor="<%=forum_bg_1%>">
    <td width="50%">&nbsp;<%=register_job%></td><td width="50%"><input type="text" name="meslek" class="form1" size="30"></td>
  </tr>
  <tr bgcolor="<%=forum_bg_2%>">
    <td width="50%">&nbsp;<%=register_sex%></td><td width="50%">
<select size="1" class="form1" name="cinsiyet">
<option value="0" selected>-------------- Se�iniz --------------</option>
<option value="a"><%=male%></option>
<option value="b"><%=female%></option>
</select>
</td>
  </tr>
  <tr bgcolor="<%=forum_bg_1%>">
    <td width="50%">&nbsp;<%=register_age%></td><td width="50%">
<select name="yas_1" size="1" class="form1">
<%
For i_y = 1 TO 31
Response.Write "<option value="""&i_y&""">"&i_y&"</option>"
Next
%>
</select>
&nbsp;/&nbsp;
<select name="yas_2" size="1" class="form1">
<%
For i_y1 = 1 TO 12
Response.Write "<option value="""&i_y1&""">"&i_y1&"</option>"
Next
%>
</select>
&nbsp;/&nbsp;
<select name="yas_3" size="1" class="form1">
<%
For i_y2 = 1940 TO Year(Now())-1
Response.Write "<option value="""&i_y2&""">"&i_y2&"</option>"
Next
%>
</select>
</td>
  </tr>
  <tr bgcolor="<%=forum_bg_1%>">
    <td width="50%">&nbsp;<%=register_signature%></td><td width="50%"><textarea name="imza" rows="5" cols="31" class="form1"></textarea></td>
  </tr>
  <tr bgcolor="<%=forum_bg_2%>">
    <td width="50%">&nbsp;<%=security_code%></td><td width="50%"><b><%=reg_sec_code%></b></td>
  </tr>
  <tr bgcolor="<%=forum_bg_1%>">
    <td width="50%">&nbsp;<%=security_code_type%></td><td width="50%"><input type="text" onkeyup="sayi(this);" name="security_code" class="form1" size="30"></td>
  </tr>
  <tr bgcolor="<%=forum_bg_2%>">
	<td colspan="2"><a ONCLICK="window.open('sozlesme.asp','peno','top=20,left=20,width=450,height=300,toolbar=no,scrollbars=yes');"><%=sett("site_isim")%> �yelik s�zle�mesini okudum ve kabul ediyorum.</a> <INPUT onclick=agreesubmit(this) type=checkbox></td>
  </tr>
  <tr bgcolor="<%=forum_bg_1%>">
	<td colspan="2" align="center"><input disabled type="submit" value="<%=uye_kayit%>" onClick="populateform(this.form.uzunluk.value)"  class="submit"></td>
  </tr>
</table>
</td>
</tr>
</table>
</form>
<%
End Sub
'#############################################################################################################################################
Sub reg
kuladi = duz(Request.Form("kuladi"))
password = duz(Request.Form("password"))
password = MN_PP(password)
email = duz(Request.Form("email"))
isim = duz(Request.Form("isim"))
g_soru = duz(Request.Form("g_soru"))
g_cevap = duz(Request.Form("g_cevap"))
g_cevap = MN_PP(g_cevap)
sehir = duz(Request.Form("sehir"))
meslek = duz(Request.Form("meslek"))
cinsiyet = duz(Request.Form("cinsiyet"))
yas = Request.Form("yas_1") & "." & Request.Form("yas_2") & "." & Request.Form("yas_3")
sc = duz(Request.Form("security_code"))
imza = duz(Request.Form("imza"))
gk = duz(Request.Form("gk"))
IF imza = "" Then
imza = "-"
ELSE
imza = imza
End IF
Set nameCheck = Connection.Execute("Select * From MEMBERS where UCase(kul_adi) = '"&UCase(kuladi)&"'")
Set mailCheck = Connection.Execute("Select * From MEMBERS where email = '"&email&"'")
IF InStr(1,kuladi,"�",1) OR InStr(1,kuladi,"�",1) OR InStr(1,kuladi,"�",1) Then
chk_kul_adi = "False"
ELSE
chk_kul_adi = "True"
End IF
If EmailCheck(email) Then
emailC = True
ELSE
emailC = False
End If
IF sc <> Session("RegSecurityCode") Then
s_code = False
ELSE
s_code = True
End IF
avatar = "IMAGES/avatars/blank.gif"
if NOT mailCheck.eof Then
response.write WriteMsg(error10)
elseif NOT nameCheck.eof Then
response.write WriteMsg(error9)
elseif emailC = False Then
response.write WriteMsg(error13)
elseif chk_kul_adi = "False" Then
response.write WriteMsg(error16)
elseif s_code = False Then
Response.Write WriteMsg(sc_err2)
ELSE
Set rec = Server.CreateObjecT("ADODB.RecordSet")
rSQL = "Select * FROM MEMBERS"
rec.open rSQL,Connection,3,3
rec.AddNew
rec("kul_adi") = kuladi
rec("sifre") = password
rec("email") = email
rec("isim") = isim
rec("g_soru") = g_soru
rec("g_cevap") = g_cevap
rec("sehir") = sehir
rec("meslek") = meslek
rec("cinsiyet") = cinsiyet
rec("yas") = yas
rec("imza") = imza
rec("login_sayisi") = 0
rec("uyelik_tarihi") = Date()
Session.LCID = 1033
rec("son_tarih") = Now()
rec("seviye") = 0
rec("msg_sayisi") = 0
rec("u_theme") = "b"
rec("u_avatar") = avatar
rec("gk") = gk
rec.Update

if sett("onaylama") = True then
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.To = email
objCDO.From = ""&sett("site_isim")&"<"&sett("site_mail")&">"
objCDO.Subject = ""&sett("site_isim")&" Uyelik Aktivitasyonu"
objCDO.BodyFormat=0
objCDO.MailFormat=0
govde=govde &" <font face=""Tahoma"" style=""font-size:11px;""> "& chr(10)
govde=govde &" Sayin <b>"&isim&"</b>; "& chr(10)
govde=govde &" <br> "& chr(10)
govde=govde &" <b>"&sett("site_isim")&"</b> isimli sitedeki �yeli�inizi aktifle�tirmek i�in l�tfen alt k�s�mdaki linki t�lay�n."& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &"<a href="""&sett("site_adres")&"/onay.asp?action=onayla&gk="&gk&""">UYELIGIMI ONAYLA</a>"& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" <b><a href="""&sett("site_adres")&""">"&sett("site_isim")&"</a></b> (Powered by <a href=""http://www.maxinuke.com"">Maxi Nuke</a>) "& chr(10)
govde=govde &"</font>"& chr(10)
objCDO.Body = govde
objCDO.Importance = 2
objCDO.Send
Set objCDO = Nothing
End if

if sett("onaylama") = false then
Response.Write WriteMsg(cong)
elseif sett("onaylama") = true then
Response.Write WriteMsg(congggg)
End if


END IF
mailCheck.Close
Set mailCheck = Nothing
nameCheck.Close
Set nameCheck = Nothing
End Sub
'#################################################################################################################################################
Sub lostpassword
step = Request.QueryString("step")
step = QS_CLEAR(step)
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
	<tr>
		<td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=uye_lostpwd%></B></CENTER></td>
	</tr>
	<tr>
		<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu"><BR><BR>
<%if step = "1" Then %>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.memname.value == "") {alert("<%=error1%>"); return false; }
return true;
}
</SCRIPT>
		<form onSubmit="return formkontrol(this)" action="?action=lostpass&step=2" method="post">
		<table border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
			<tr>
				<td align="right"><%=lost_memname%> : </td>
				<td><input type="text" name="memname" size="30" class="form1" maxlength="15"></td>
			</tr>
			<tr>
				<td colspan="2"><BR><BR></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><input type="submit" value="<%=lost_continue%>" class="submit"></td>
			</tr>
		</table>
		</form>
<%
ElseIf step = "2" Then
member = Request.Form("memname")
Set uyectrl = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&member&"'")
	IF member<>"" Then
		IF uyectrl.eof Then
		Response.Write error12
		ELSE
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.answer.value == "") {alert("<%=error6%>"); return false; }
return true;
}
</SCRIPT>
		<form onSubmit="return formkontrol(this)" action="?action=lostpass&step=3&memname=<%=uyectrl("kul_adi")%>" method="post">
		<table border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
			<tr>
				<td colspan="2" align="center"><%=uyectrl("g_soru")%></td>
			</tr>
			<tr>
				<td align="right"><%=register_answer%> : </td>
				<td><input type="text" name="answer" size="30" class="form1"></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><input type="submit" value="<%=lost_continue%>" class="submit"></td>
			</tr>
		</table>
		</form>
<%
		End If
	End If
uyectrl.Close
Set uyectrl = Nothing
ElseIf step = "3" Then
answer = Request.Form("answer")
member = Request.QueryString("memname")
member = QS_CLEAR(member)
Set answerctrl = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&member&"'")
	IF answerctrl("g_cevap") <> MN_PP(answer) Then
	response.write error11
	Else
		IF sett("fpass") = "Site" Then
		asd = answerctrl("sifre")
		mem_new_pass = CPASS(16)
		Connection.Execute("UPDATE MEMBERS SET sifre = '"&MN_PP(mem_new_pass)&"' WHERE kul_adi = '"&member&"'")
%>
		<table border="0" cellspacing="0" cellpadding="0" align="center" width="75%" class="td_menu">
			<tr>
				<td colspan="2" align="center"><%=member%></td>
			</tr>
			<tr>
				<td width="50%" align="right"><%=uye_new_password%> : </td>
				<td width="50%"><b><%=mem_new_pass%></b></td>
			</tr>
		</table>
<%
		ElseIF sett("fpass") = "Mail" Then
		Set objCDO = Server.CreateObject("CDONTS.NewMail")
		objCDO.To = answerctrl("email")
		objCDO.From = sett("site_mail")
		objCDO.Subject = ""&sett("site_isim")&" �ifre Hat�rlatma"
		objCDO.BodyFormat=0
		objCDO.MailFormat=0
		govde=govde &" <font face=""Tahoma"" style=""font-size:11px;""> "& chr(10)
		govde=govde &" Sayin <b>"&answerctrl("kul_adi")&"</b>; "& chr(10)
		govde=govde &" <br> "& chr(10)
		govde=govde &" <b>"&sett("site_isim")&"</b> isimli sitedeki Sifre Hatirlatma servisi kullanilarak sifreniz mailinize iletilmistir.Mailin sizinle bir alakasi olmadigini d�s�n�yorsaniz bu maili siliniz... "& chr(10)
		govde=govde &" <br><br> "& chr(10)
		govde=govde &" Yeni Sifreniz : <b>"&mem_new_pass&"</b> "& chr(10)
		govde=govde &" <br><br> "& chr(10)
		govde=govde &" <b><a href="""&sett("site_adres")&""">"&sett("site_isim")&"</a></b> (Powered by <a href=""http://www.maxinuke.com"">Maxi Nuke</a>) "& chr(10)
		govde=govde &"</font>"& chr(10)
		objCDO.Body = govde
		objCDO.Importance = 2
		objCDO.Send
		Set objCDO = Nothing
		Response.Write fp_msg
		End If
	End If
answerctrl.Close
Set answerctrl = Nothing
End If
%>
<br><br><div align=center class=td_menu><a href="javascript:history.back();"><%=p_back%></div>
		</td>
	</tr>
</table>
<%
End Sub
call ORTA2
call ALT
%>