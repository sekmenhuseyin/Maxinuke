<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
If act = "members" Then
call members
ElseIf act = "member_delete" Then
call mem_del
ElseIf act = "member_details" Then
call mem_details
ElseIf act = "member_update" Then
call mem_update
ElseIf act = "AdminMsgs" Then
call adminmsgs
ElseIf act = "MsgToMembers" Then
call msgmembers
ElseIf act = "SendMsg" Then
call sendmsg
ElseIf act = "MemMsgRead" Then
call read_memmsg
ElseIf act = "DelMemMsg" Then
call delete_memmsg
ElseIF act = "Mail2Members" Then
call mail_2_members
ElseIF act = "SendMail2Members" Then
call mail_send_2_members
ElseIf act = "MemberMsgs" Then
call membermsgs
ElseIF act = "DeleteMessages" Then
call del_all_msgs
ElseIF act = "onaysizsil" Then
call onaysiz_sil
End If

Sub members
%>
<script language="JavaScript">
<!-- // silim islemi onay uyarisi
function submitConfirm18(listForm2)
{ 
   listForm2.target="_self"; 
   listForm2.action="";
   var answer = confirm ("Bu �yeyinin silinmesini onayl�yor musunuz ?") 
   if (answer)
       return true;
   else
       return false;
} 
//-->
</script>
<%IF Len(Request.QueryString("y")) < "1" Then
arrange_word = "A"
ELSE
arrange_word = Request.QueryString("y")
End IF
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MEMBERS WHERE Left(kul_adi,1) LIKE '%"&arrange_word&"%' ORDER BY kul_adi ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="26" align="center">-=- �YE L�STES� -=-</td>
	</tr>
	<tr>
<%For b = 65 To 90 Step 1%>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"><%= "<a href='?action=members&y=" & Chr(b) &"'>" & Chr(b) &"</a>"%></td>
<%Next%>
	</tr>
</table>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<%While Not rs.EOF%>
  <tr>
    <td><a href="?action=member_details&mid=<%=rs("uye_id")%>"><%=rs("kul_adi")%></a>[<a href="?action=member_delete&id=<%=rs("uye_id")%>" onClick="return submitConfirm18(this)">Sil</a>]</td>
<%
rs.MoveNext
If Not rs.EOF Then
%>
    <td><a href="?action=member_details&mid=<%=rs("uye_id")%>"><%=rs("kul_adi")%></a>[<a href="?action=member_delete&id=<%=rs("uye_id")%>" onClick="return submitConfirm18(this)">Sil</a>]</td>
<%
End If
If Not rs.EOF Then
rs.MoveNext
End If
%>
</tr>
<% Wend %>
</table>
<%
End Sub
Sub mem_del
uid = Request.QueryString("id")
IF uid<>"" Then
Set dm = Server.CreateObject("ADODB.RecordSet")
dmSQL = "DELETE * FROM MEMBERS WHERE uye_id = "&uid&""
dm.open dmSQL,Connection,1,3
End If
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub mem_details
uid = Request.QueryString("mid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MEMBERS WHERE uye_id = "&uid&""
rs.open SQL,Connection,1,3
uye_gcevap = rs("g_cevap")
age_check = Split(rs("yas"),".")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.membername.value == "") {alert("<%=error1%>"); return false; }
if (form.email.value == "") {alert("<%=error3%>"); return false; }
if (form.isim.value == "") {alert("<%=error4%>"); return false; }
if (form.g_soru.value == "") {alert("<%=error5%>"); return false; }
if (form.sehir.value == "0") {alert("L�tfen Uyenin Yasadigi Sehri Belirtiniz"); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<form onSubmit="return formkontrol(this)" action="?action=member_update&mid=<%=uid%>" method="post">
	<tr>
		<td colspan="7" align="center"><B>-=- <%=rs("kul_adi")%> TAKMA ADLI UYENIN BILGILERI -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">�ye Ad�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Yeni �ifre</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">E-Mail</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ad Soyad</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Gizli Soru</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Gizli Cevap</td>
	</tr>
	<tr>
		<td align="center"><input type="text" name="membername" class="form1" value="<%=rs("kul_adi")%>"></td>
		<td align="center"><input type="text" name="password" class="form1"></td>
		<td align="center"><input type="text" name="email" class="form1" value="<%=rs("email")%>"></td>
		<td align="center"><input type="text" name="isim" class="form1" value="<%=rs("isim")%>"></td>
		<td align="center"><input type="text" name="g_soru" class="form1" value="<%=rs("g_soru")%>"></td>
		<td align="center"><input type="text" name="g_cevap" class="form1"></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">�ehir</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Meslek</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Cinsiyet</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ya�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Giris Say�s�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">�yelik Tarihi</td>
	</tr>
	<tr>
		<td align="center">
		<%=rs("sehir")%>
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
		<td align="center"><input type="text" name="meslek" class="form1" value="<%=rs("meslek")%>"></td>
		<td align="center">
		<%
		if rs("cinsiyet") = "a" then
		section1 = "selected"
		elseif rs("cinsiyet") = "b" then
		section2 = "selected"
		else
		section1 = "selected"
		end if
		%>
		<select size="1" class="form1" name="cinsiyet">
		<option value="a" <%=section1%>>Bay</option>
		<option value="b" <%=section2%>>Bayan</option>
		</select>
		</td>
		<td align="center">
		<select name="yas_1" size="1" class="form1">
		<%
		For i_y = 1 TO 31
		IF Len(i_y) = "1" Then
		i_y = "0" & i_y
		End IF
		IF Left(i_y,2) = age_check(0) Then
		y_opt = "selected"
		ELSE
		y_opt = ""
		End IF
		Response.Write "<option value="""&i_y&""" "&y_opt&">"&i_y&"</option>"
		Next
		%>
		</select>
		.
		<select name="yas_2" size="1" class="form1">
		<%
		For i_y1 = 1 TO 12
		IF Len(i_y1) = "1" Then
		i_y1 = "0" & i_y1
		End IF
		IF Left(i_y1,2) = age_check(1) Then
		y_opt1 = "selected"
		ELSE
		y_opt1 = ""
		End IF
		Response.Write "<option value="""&i_y1&""" "&y_opt1&">"&i_y1&"</option>"
		Next
		%>
		</select>
		.
		<select name="yas_3" size="1" class="form1">
		<%
		For i_y2 = 1920 TO Year(Now())-1
		IF Left(i_y2,4) = age_check(2) Then
		y_opt2 = "selected"
		ELSE
		y_opt2 = ""
		End IF
		Response.Write "<option value="""&i_y2&""" "&y_opt2&">"&i_y2&"</option>"
		Next
		%>
		</select>
		</td>
		<td align="center"><%=rs("login_sayisi")%></td>
		<td align="center"><%=rs("uyelik_tarihi")%></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Son Geli�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Avatar�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Mesaj Say�s�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">�p Adresi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Temas�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Seviye</td>
	</tr>
	<tr>
		<td align="center"><%=rs("son_tarih")%></td>
		<td align="center"><IMG SRC="<%'=rs("u_avatar")%>"></td>
		<td align="center"><input type="text" name="msg" class="form1" value="<%=rs("msg_sayisi")%>"></td>
		<td align="center"><%=rs("last_ip")%></td>
		<td align="center"><%=rs("u_theme")%></td>
		<td align="center">
		<%
		If rs("seviye") = "0" Then
		sel = "selected"
		ElseIf rs("seviye") = "1" Then
		sel1 = "selected"
		ElseIf rs("seviye") = "2" Then
		sel2 = "selected"
		ElseIf rs("seviye") = "3" Then
		sel3 = "selected"
		ElseIf rs("seviye") = "4" Then
		sel4 = "selected"
		ElseIf rs("seviye") = "5" Then
		sel5 = "selected"
		End If
		%>
		<select size="1" name="seviye" class="form1">
		<option value="0" <%=sel%>>�ye</option>
		<option value="1" <%=sel1%>>Y�netici</option>
		<option value="2" <%=sel2%>>Download Edit�r�</option>
		<option value="3" <%=sel3%>>Makale Edit�r�</option>
		<option value="4" <%=sel4%>>Haber Edit�r�</option>
		<option value="5" <%=sel5%>>Forum G�revlisi</option>
		</select>		
		</td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Browser�</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">��letim Sistemi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Onl�ne m� ?</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"></td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"></td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"></td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif"></td>
	</tr>
	<tr>
		<td align="center"><%=rs("u_browser")%></td>
		<td align="center"><%=rs("u_ws")%></td>
		<td align="center">
		<%
		if rs("u_online") = True Then
		response.write "Evet"
		else
		response.write "Hayir"
		End if
		%></td>
		<td></td>
		<td></td>
		<td></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif" colspan="7">Imzasi</td>
	</tr>
	<tr>
		<td align="center" colspan="7"><textarea name="imza" rows="10" COLS="100" class="form1"><%=rs("imza")%></textarea></td>
	</tr>
	<tr>
		<td colspan="7"><input type="submit" value="De�i�iklikleri Kaydet" class="form1" style="width:100%"></td>
  </tr>
</form>
</table>
<%
End Sub

Sub mem_update
uid = Request.QueryString("mid")
memname = duz(Request.Form("membername"))
password = duz(Request.Form("password"))
email = duz(Request.Form("email"))
isim = duz(Request.Form("isim"))
g_soru = duz(Request.Form("g_soru"))
g_cevap = duz(Request.Form("g_cevap"))
sehir = duz(Request.Form("sehir"))
meslek = duz(Request.Form("meslek"))
cinsiyet = duz(Request.Form("cinsiyet"))
yas = Request.Form("yas_1") & "." & Request.Form("yas_2") & "." & Request.Form("yas_3")
logincount = duz(Request.Form("login"))
msgcount = duz(Request.Form("msg"))
level = duz(Request.Form("seviye"))
imza = duz(Request.Form("imza"))

IF Len(imza) <= 0 Then
imza = "-"
ELSE
imza = imza
End IF

if memname="" then
response.write "<center>L�tfen �ye Ad� Yaz�n�z</center>"

elseif email="" then
response.write "<center>L�tfen Mail Adresi Yaz�n�z</center>"

elseif isim="" then
response.write "<center>L�tfen �sim Yaz�n�z</center>"

elseif g_soru="" then
response.write "<center>L�tfen Gizli Soru Yaz�n�z</center>"

ELSE

Set rec = Server.CreateObjecT("ADODB.RecordSet")
rSQL = "Select * FROM MEMBERS WHERE uye_id = "&uid&""
rec.open rSQL,Connection,1,3

IF password<>"" Then
password = MN_PP(password)
ELSE
password = rec("sifre")
End IF

IF g_cevap<>"" Then
g_cevap = MN_PP(g_cevap)
ELSE
g_cevap = rec("g_cevap")
End IF

rec("kul_adi") = memname
rec("sifre") = password
rec("email") = email
rec("isim") = isim
rec("g_soru") = g_soru
rec("g_cevap") = g_cevap
rec("sehir") = sehir
rec("meslek") = meslek
rec("cinsiyet") = cinsiyet
rec("yas") = yas
rec("msg_sayisi") = msgcount
rec("seviye") = level
rec("imza") = imza
rec.Update

Response.Write "<BR><BR><BR><div class=td_menu align=center>Bilgiler Ba�ar�yla G�ncellendi...</div>"
END IF
End Sub

Sub adminmsgs
Set rs = Connection.Execute("Select * FROM MESSAGES WHERE TYPE = 1")
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="3" align="center"><B>-=- TOPLU MESAJ -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Konu</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% do while not rs.eof %>
<tr>
<td><a href="?action=MemMsgRead&mid=<%=rs("mid")%>"><%=rs("subject")%></a></td>
<td align="center"><%=rs("mdate")%></td>
<td align="center"><a href="?action=DelMemMsg&mid=<%=rs("mid")%>"><img src="images/temalar/<%=Session("SiteTheme")%>/msg_del.gif" border="0" alt="Bu Mesaj� Tamamen Silmek i�in T�klay�n !!!"></a></td>
</tr>
<%
rs.MoveNext
Loop

rs.Close
Set rs = Nothing
End Sub


Sub msgmembers
%>
<form action="?action=SendMsg" method="post">
<table border="0" cellspacing="2" cellpadding="1" width="90%" align="center" class="td_menu">
<tr >
<td width="20%" align="right">&nbsp;Konu :&nbsp;</td><td><input type="text" name="subject" class="form1" style="width:100%"></td>
</tr>
<tr >
<td colspan="2">
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.Value	= ""
oFCKeditor.Create "msg"
%>
</td>
</tr>
<tr>
<td colspan="2" align="right"><input type="submit" value="   G�nder   " class="submit" style="width:100%"></td>
</tr> 
</table>
</form>
<%
End Sub

Sub sendmsg
subj = Request.Form("subject")
message = Request.Form("msg")
from = "Site Y�netimi"

If subj="" Then
Response.Write "L�tfen Konu Yaz�n�z..."
ElseIf message="" Then
Response.Write "L�tfen Mesaj Yaz�n�z..."
ELSE

Set enter = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM MESSAGES"
enter.open eSQL,Connection,1,3

enter.AddNew
enter("from") = from
enter("msg") = message
enter("mdate") = Now()
enter("read") = False
enter("subject") = subj
enter("type") = 1
enter("see") = True
enter.Update
Response.Write "<BR><BR><BR><BR><BR><CENTER><div class=td_menu>Mesaj T�m �yelere �letildi...</div></CENTER>"
END IF
End Sub

Sub read_memmsg
id = Request.QueryString("mid")
Set rs = Connection.Execute("Select * FROM MESSAGES WHERE mid = "&id&"")
If rs("type") = "0" Then
tto = rs("to")
ElseIf rs("type") = "1" Then
tto = "T�m �yelere"
Else
tto = ""
End If
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="6" align="center"><B>-=- GONDERILMIS TOPLU MESAJ -=-</B></td>
	</tr>

<tr>
<td align="right" >Kimden : </td>
<td ><%=rs("from")%></td>
<td align="right" >Kime : </td>
<td ><%=tto%></td>
<td align="right" >Tarih : </td>
<td ><%=rs("mdate")%></td>
</tr>
<tr>
<td colspan="6"><CENTER>Mesaj</CENTER><br><%=rs("msg")%></td>
</tr>
<tr>
<td colspan="6"><center><a href="?action=DelMemMsg&mid=<%=id%>">Bu Mesaj� Tamamen Sil</a></center></td>
</tr>
</table>
<%
rs.Close
Set rs = Nothing
End Sub

Sub delete_memmsg
id = Request.QueryString("mid")
Set del = Server.CreateObject("ADODB.RecordSet")
dSQL = "DELETE * FROM MESSAGES WHERE mid = "&id&""
del.open dSQL,Connection,1,3
Response.Write "<div class=td_menu><BR><BR><BR><center>Mesaj Ba�ar�yla Silindi...</center></div>"
End Sub

Sub mail_2_members
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" width="90%" align="center" class="td_menu">
<form method="post" action="?action=SendMail2Members">
<tr>
<td width="20%" align="right">Ba�l�k : </td>
<td width="80%"><input type="text" name="m_title" class="form1" style="width:100%"></td>
</tr>
<tr>
<td colspan="2">
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.Value	= ""
oFCKeditor.Create "m_msg"
%></td>
</tr>
<tr align="center">
<td colspan="2"><input type="submit" value="G�nder" class="submit" style="width:100%"></td>
</tr>
</form>
</table>
<%
End SUB

Sub mail_send_2_members
IF Request.Form("m_title") = "" Then
Response.Write WriteMsg("L�tfen Basliginizi Yaz�n�z")
elseIF Request.Form("m_msg") = "" Then
Response.Write WriteMsg("L�tfen Mesaj�n�z� Yaz�n�z")
ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MEMBERS ORDER BY uye_id DESC",Connection,3,3
Set objSettings = Server.CreateObject("ADODB.RecordSet")
objSettings.open "Select * FROM SETTINGS",Connection,3,3
Do Until rs.EoF
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" "&Request.Form("m_title")&""
Avanos.Body="<div class=td_menu>"&Request.Form("m_msg")&"</div>"
Avanos.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
Avanos.To=""&rs("email")&""
Avanos.Importance = 2
Avanos.Send
set HTML = nothing
set Avanos=nothing
rs.MoveNext
Loop
Response.Write WriteMsg("TEBR�KLER !<br><br>Mesaj�n�z T�m �yelerin Mail Adreslerine G�nderildi.</div>")
End IF
End SUB

Sub membermsgs
Set rs = Connection.Execute("Select * FROM MESSAGES WHERE TYPE = 0")
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="6" align="center"><B>-=- �� MESAJLA�MALAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kimden</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kime</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Konu</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">G�nderim Tarihi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Okunma Tarihi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Durum</td>
	</tr>
<%
Do While NOT rs.Eof

If rs("see") = True Then
durum = "G�steriliyor"
ElseIf rs("see") = False Then
durum = "Silinmi�"
Else
durum = "Belirsiz"
End If
%>
<tr>
<td align="center"><%=rs("from")%></td>
<td align="center"><%=rs("to")%></td>
<td align="center"><a href="?action=MemMsgRead&mid=<%=rs("mid")%>"><%=rs("subject")%></a></td>
<td align="center"><%=rs("mdate")%></td>
<td align="center">.</td>
<td align="center"><%=durum%></td>
</tr>
<%
rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub


Sub del_all_msgs
IF Request.QueryString("x") = "OK" Then
Connection.Execute("DELETE FROM MESSAGES")
Response.Write "<div align=center class=td_menu>Tebrikler !!!<br><BR><BR><BR>T�m Mesajlar Ba�ar�yla Silindi...</div>"
ELSE
Response.Write "<div align=center class=td_menu>D�KKAT !!!<br><br>Tamam'a bast���n�z andan itibaret Toplu Mesajlar ve �yeler Aras�ndaki mesajlar�n hepsi silinecektir...<br><br><input type=""button"" value=""Tamam"" class=""submit"" onClick=""location.href('?action="&act&"&x=OK')""> <input type=""button"" value=""Vazge�"" class=""submit"" onClick=""location.href('default.asp?pane=Main')"">"
End IF
End Sub


Sub onaysiz_sil
Connection.Execute("DELETE FROM MEMBERS where onay=false")
Response.Write "<div align=center class=td_menu>Tebrikler !!!<br><BR><BR><BR>T�m Onaysiz uyeleriniz Ba�ar�yla Silindi...</div>"
End sub
End If
%>