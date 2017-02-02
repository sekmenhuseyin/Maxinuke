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
   var answer = confirm ("Bu üyeyinin silinmesini onaylýyor musunuz ?") 
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
		<td colspan="26" align="center">-=- ÜYE LÝSTESÝ -=-</td>
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
if (form.sehir.value == "0") {alert("Lütfen Uyenin Yasadigi Sehri Belirtiniz"); return false; }
return true;
}
</SCRIPT>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<form onSubmit="return formkontrol(this)" action="?action=member_update&mid=<%=uid%>" method="post">
	<tr>
		<td colspan="7" align="center"><B>-=- <%=rs("kul_adi")%> TAKMA ADLI UYENIN BILGILERI -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Üye Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Yeni Þifre</td>
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
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Þehir</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Meslek</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Cinsiyet</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Yaþ</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Giris Sayýsý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Üyelik Tarihi</td>
	</tr>
	<tr>
		<td align="center">
		<%=rs("sehir")%>
		<SELECT name="sehir" class="form1" >
		<OPTION value=0 selected>- Lütfen Þehir Seçimi Yapýnýz. -</OPTION>
		<OPTION value="Adana">Adana</OPTION>
		<OPTION value="Adýyaman">Adýyaman</OPTION>
		<OPTION value="Afyon">Afyon</OPTION>
		<OPTION value="Aðrý">Aðrý</OPTION>
		<OPTION value="Aksaray">Aksaray</OPTION>
		<OPTION value="Amasya">Amasya</OPTION>
		<OPTION value="Ankara">Ankara</OPTION>
		<OPTION value="Antalya">Antalya</OPTION>
		<OPTION value="Ardahan">Ardahan</OPTION>
		<OPTION value="Artvin">Artvin</OPTION>
		<OPTION value="Aydýn">Aydýn</OPTION>
		<OPTION value="Balýkesir">Balýkesir</OPTION>
		<OPTION value="Bartýn">Bartýn</OPTION>
		<OPTION value="Batman">Batman</OPTION>
		<OPTION value="Bayburt">Bayburt</OPTION>
		<OPTION value="Bilecik">Bilecik</OPTION>
		<OPTION value="Bingöl">Bingöl</OPTION>
		<OPTION value="Bitlis">Bitlis</OPTION>
		<OPTION value="Bolu">Bolu</OPTION>
		<OPTION value="Burdur">Burdur</OPTION>
		<OPTION value="Bursa">Bursa</OPTION>
		<OPTION value="Çanakkale">Çanakkale</OPTION>
		<OPTION value="Çankýrý">Çankýrý</OPTION>
		<OPTION value="Çorum">Çorum</OPTION>
		<OPTION value="Denizli">Denizli</OPTION>
		<OPTION value="Diyarbakýr">Diyarbakýr</OPTION>
		<OPTION value="Düzce">Düzce</OPTION>
		<OPTION value="Edirne">Edirne</OPTION>
		<OPTION value="Elazýð">Elazýð</OPTION>
		<OPTION value="Erzincan">Erzincan</OPTION>
		<OPTION value="Erzurum">Erzurum</OPTION>
		<OPTION value="Eskiþehir">Eskiþehir</OPTION>
		<OPTION value="Gaziantep">Gaziantep</OPTION>
		<OPTION value="Gazimagosa">Gazimagosa</OPTION>
		<OPTION value="Giresun">Giresun</OPTION>
		<OPTION value="Girne">Girne</OPTION>
		<OPTION value="Gümüþhane">Gümüþhane</OPTION>
		<OPTION value="Güzelyurt">Güzelyurt</OPTION>
		<OPTION value="Hakkari">Hakkari</OPTION>
		<OPTION value="Hatay">Hatay</OPTION>
		<OPTION value="Iðdýr">Iðdýr</OPTION>
		<OPTION value="Isparta">Isparta</OPTION>
		<OPTION value="Mersin">Ýçel (Mersin)</OPTION>
		<OPTION value="Ýskele">Ýskele</OPTION>
		<OPTION value="Ýstanbul-Anadolu">Ýstanbul-Anadolu</OPTION>
		<OPTION value="Ýstanbul-Avrupa">Ýstanbul-Avrupa</OPTION>
		<OPTION value="Ýzmir">Ýzmir</OPTION>
		<OPTION value="Kahramanmaraþ">Kahramanmaraþ</OPTION>
		<OPTION value="Karabük">Karabük</OPTION>
		<OPTION value="Karaman">Karaman</OPTION>
		<OPTION value="Kars">Kars</OPTION>
		<OPTION value="Kastamonu">Kastamonu</OPTION>
		<OPTION value="Kayseri">Kayseri</OPTION>
		<OPTION value="Kýrýkkale">Kýrýkkale</OPTION>
		<OPTION value="Kýrklareli">Kýrklareli</OPTION>
		<OPTION value="Kýrþehir">Kýrþehir</OPTION>
		<OPTION value="Kilis">Kilis</OPTION>
		<OPTION value="Kocaeli">Kocaeli</OPTION>
		<OPTION value="Konya">Konya</OPTION>
		<OPTION value="Kütahya">Kütahya</OPTION>
		<OPTION value="Lefkosa">Lefkosa</OPTION>
		<OPTION value="Malatya">Malatya</OPTION>
		<OPTION value="Manisa">Manisa</OPTION>
		<OPTION value="Mardin">Mardin</OPTION>
		<OPTION value="Muðla">Muðla</OPTION>
		<OPTION value="Muþ">Muþ</OPTION>
		<OPTION value="Nevþehir">Nevþehir</OPTION>
		<OPTION value="Niðde">Niðde</OPTION>
		<OPTION value="Ordu">Ordu</OPTION>
		<OPTION value="Osmaniye">Osmaniye</OPTION>
		<OPTION value="Rize">Rize</OPTION>
		<OPTION value="Sakarya">Sakarya</OPTION>
		<OPTION value="Samsun">Samsun</OPTION>
		<OPTION value="Siirt">Siirt</OPTION>
		<OPTION value="Sinop">Sinop</OPTION>
		<OPTION value="Sivas">Sivas</OPTION>
		<OPTION value="Þanlýurfa">Þanlýurfa</OPTION>
		<OPTION value="Þýrnak">Þýrnak</OPTION>
		<OPTION value="Tekirdað">Tekirdað</OPTION>
		<OPTION value="Tokat">Tokat</OPTION>
		<OPTION value="Trabzon">Trabzon</OPTION>
		<OPTION value="Tunceli">Tunceli</OPTION>
		<OPTION value="Uþak">Uþak</OPTION>
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
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Son Geliþ</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Avatarý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Mesaj Sayýsý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýp Adresi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Temasý</td>
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
		<option value="0" <%=sel%>>Üye</option>
		<option value="1" <%=sel1%>>Yönetici</option>
		<option value="2" <%=sel2%>>Download Editörü</option>
		<option value="3" <%=sel3%>>Makale Editörü</option>
		<option value="4" <%=sel4%>>Haber Editörü</option>
		<option value="5" <%=sel5%>>Forum Görevlisi</option>
		</select>		
		</td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Browserý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýþletim Sistemi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Onlýne mý ?</td>
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
		<td colspan="7"><input type="submit" value="Deðiþiklikleri Kaydet" class="form1" style="width:100%"></td>
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
response.write "<center>Lütfen Üye Adý Yazýnýz</center>"

elseif email="" then
response.write "<center>Lütfen Mail Adresi Yazýnýz</center>"

elseif isim="" then
response.write "<center>Lütfen Ýsim Yazýnýz</center>"

elseif g_soru="" then
response.write "<center>Lütfen Gizli Soru Yazýnýz</center>"

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

Response.Write "<BR><BR><BR><div class=td_menu align=center>Bilgiler Baþarýyla Güncellendi...</div>"
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
<td align="center"><a href="?action=DelMemMsg&mid=<%=rs("mid")%>"><img src="images/temalar/<%=Session("SiteTheme")%>/msg_del.gif" border="0" alt="Bu Mesajý Tamamen Silmek için Týklayýn !!!"></a></td>
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
<td colspan="2" align="right"><input type="submit" value="   Gönder   " class="submit" style="width:100%"></td>
</tr> 
</table>
</form>
<%
End Sub

Sub sendmsg
subj = Request.Form("subject")
message = Request.Form("msg")
from = "Site Yönetimi"

If subj="" Then
Response.Write "Lütfen Konu Yazýnýz..."
ElseIf message="" Then
Response.Write "Lütfen Mesaj Yazýnýz..."
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
Response.Write "<BR><BR><BR><BR><BR><CENTER><div class=td_menu>Mesaj Tüm Üyelere Ýletildi...</div></CENTER>"
END IF
End Sub

Sub read_memmsg
id = Request.QueryString("mid")
Set rs = Connection.Execute("Select * FROM MESSAGES WHERE mid = "&id&"")
If rs("type") = "0" Then
tto = rs("to")
ElseIf rs("type") = "1" Then
tto = "Tüm Üyelere"
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
<td colspan="6"><center><a href="?action=DelMemMsg&mid=<%=id%>">Bu Mesajý Tamamen Sil</a></center></td>
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
Response.Write "<div class=td_menu><BR><BR><BR><center>Mesaj Baþarýyla Silindi...</center></div>"
End Sub

Sub mail_2_members
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" width="90%" align="center" class="td_menu">
<form method="post" action="?action=SendMail2Members">
<tr>
<td width="20%" align="right">Baþlýk : </td>
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
<td colspan="2"><input type="submit" value="Gönder" class="submit" style="width:100%"></td>
</tr>
</form>
</table>
<%
End SUB

Sub mail_send_2_members
IF Request.Form("m_title") = "" Then
Response.Write WriteMsg("Lütfen Basliginizi Yazýnýz")
elseIF Request.Form("m_msg") = "" Then
Response.Write WriteMsg("Lütfen Mesajýnýzý Yazýnýz")
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
Response.Write WriteMsg("TEBRÝKLER !<br><br>Mesajýnýz Tüm Üyelerin Mail Adreslerine Gönderildi.</div>")
End IF
End SUB

Sub membermsgs
Set rs = Connection.Execute("Select * FROM MESSAGES WHERE TYPE = 0")
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="6" align="center"><B>-=- ÝÇ MESAJLAÞMALAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kimden</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kime</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Konu</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Gönderim Tarihi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Okunma Tarihi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Durum</td>
	</tr>
<%
Do While NOT rs.Eof

If rs("see") = True Then
durum = "Gösteriliyor"
ElseIf rs("see") = False Then
durum = "Silinmiþ"
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
Response.Write "<div align=center class=td_menu>Tebrikler !!!<br><BR><BR><BR>Tüm Mesajlar Baþarýyla Silindi...</div>"
ELSE
Response.Write "<div align=center class=td_menu>DÝKKAT !!!<br><br>Tamam'a bastýðýnýz andan itibaret Toplu Mesajlar ve Üyeler Arasýndaki mesajlarýn hepsi silinecektir...<br><br><input type=""button"" value=""Tamam"" class=""submit"" onClick=""location.href('?action="&act&"&x=OK')""> <input type=""button"" value=""Vazgeç"" class=""submit"" onClick=""location.href('default.asp?pane=Main')"">"
End IF
End Sub


Sub onaysiz_sil
Connection.Execute("DELETE FROM MEMBERS where onay=false")
Response.Write "<div align=center class=td_menu>Tebrikler !!!<br><BR><BR><BR>Tüm Onaysiz uyeleriniz Baþarýyla Silindi...</div>"
End sub
End If
%>