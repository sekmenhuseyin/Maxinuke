<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA
IF Session("Enter") = "Yes" Then
%>
                <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><CENTER><%=ya_topic%></CENTER></B></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<%
	IF QS_CLEAR(Request.QueryString("op")) = "UpdateProfile" Then

email = duz(Request.Form("email"))
isim = duz(Request.Form("isim"))
g_soru = duz(Request.Form("g_soru"))
g_cevap = duz(Request.Form("g_cevap"))
IF g_cevap <> "" Then
g_cevap = MN_PP(g_cevap)
End IF
msn = duz(Request.Form("msn"))
sehir = duz(Request.Form("sehir"))
meslek = duz(Request.Form("meslek"))
cinsiyet = duz(Request.Form("cinsiyet"))
yas = Request.Form("yas_1") & "." & Request.Form("yas_2") & "." & Request.Form("yas_3")
signimza = duz(Request.Form("imza"))


IF signimza = "" Then
signimza = "-"
ELSE
signimza = signimza
End IF

IF isim="" then
response.write "<b>" & error4 & "</b>"

ElseIF g_soru="" then
response.write "<b>" & error5 & "</b>"

	If sett("onaylama")=false Then
ElseIF EmailCheck(email) = False Then
Response.Write "<b>" & error13 & "</b>"
	End if

ELSE

If sett("onaylama")=false Then
Connection.Execute("UPDATE MEMBERS SET email = '"&email&"' WHERE uye_id = "&Session("Uye_ID")&"")
End if
Connection.Execute("UPDATE MEMBERS SET isim = '"&isim&"' WHERE uye_id = "&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET g_soru = '"&g_soru&"' WHERE uye_id = "&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET sehir = '"&sehir&"' WHERE uye_id = "&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET meslek = '"&meslek&"' WHERE uye_id = "&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET cinsiyet = '"&cinsiyet&"' WHERE uye_id = "&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET yas = '"&yas&"' WHERE uye_id = "&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET imza = '"&signimza&"' WHERE uye_id = "&Session("Uye_ID")&"")

IF g_cevap<>"" Then
Connection.Execute("UPDATE MEMBERS SET g_cevap = '"&g_cevap&"' WHERE uye_id = "&Session("Uye_ID")&"")
End IF

Response.Write cong2

END IF

	ElseIF QS_CLEAR(Request.QueryString("op")) = "Profile" Then

uid = Session("Uye_ID")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MEMBERS WHERE uye_id = "&uid&""
rs.open SQL,Connection,3,3

age_check = Split(rs("yas"),".")
%>
<table width="100%" border="0" cellspacing="2" cellpadding="1" align="center" bgcolor="<%=content_bg%>" class="td_menu" style="font-weight:bold">
<form action="?id=<%=uid%>&op=UpdateProfile" method="post">
  <tr height="20">
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=uye_kuladi%></td><td bgcolor="<%=content_bg%>" width="50%"><b><%=Session("Member")%></b></td>
  </tr>
  <tr height="20">
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=uye_password%></td><td bgcolor="<%=content_bg%>" width="50%"><b>[<a href="?op=ChangePass"><%=pass_change%></a>]</td>
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=register_mail%></td><td bgcolor="<%=content_bg%>" width="50%"><% If sett("onaylama")=True then %><%=rs("email")%><%else%><input type="text" name="email" class="form1" value="<%=rs("email")%>" style="width:100%" size="20"><%End if%></td>
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;Ad ve Soyadýnýz</td><td bgcolor="<%=content_bg%>" width="50%">
    <input type="text" name="isim" class="form1" value="<%=rs("isim")%>" style="width:100%" size="20"></td>
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=register_question%></td><td bgcolor="<%=content_bg%>" width="50%">
    <input type="text" name="g_soru" class="form1" value="<%=rs("g_soru")%>" style="width:100%" size="20"></td>
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=register_answer%><br><font face="verdana" size="1" color="gray">(<%=gc_information%>)</font></td><td bgcolor="<%=content_bg%>" width="50%">
    <input type="text" name="g_cevap" class="form1" style="width:100%" size="20"></td>
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=register_city%></td><td bgcolor="<%=content_bg%>" width="50%">
<%'=rs("sehir")%>


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
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=register_job%></td><td bgcolor="<%=content_bg%>" width="50%">
    <input type="text" name="meslek" class="form1" value="<%=rs("meslek")%>" style="width:100%" size="20"></td>
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;<%=register_sex%></td><td bgcolor="<%=content_bg%>" width="50%">
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
<option value="a" <%=section1%>><%=male%></option>
<option value="b" <%=section2%>><%=female%></option>
</select>
</td>
  </tr>
  <tr> 
    <td bgcolor="<%=content_bg%>" width="50%">&nbsp;Dogum Tarihi</td><td bgcolor="<%=content_bg%>" width="50%">
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
  </tr>
 <tr> 
    <td bgcolor="<%=content_bg%>" width="50%" valign="top">&nbsp;<%=register_signature%></td><td bgcolor="<%=content_bg%>" width="50%">
    <textarea rows="5" name="imza" class="form1" style="width:100%" cols="20"><% IF rs("imza") <> "" Then %><%=Oku(rs("imza"))%><% End IF %></textarea></td>
  </tr>
<tr> 
    <td colspan="2" bgcolor="<%=content_bg%>"><input type="submit" value="<%=uye_guncelle%>" class="submit" style="width:100%"></td>
  </tr>
</form>
</table>
<%
	ElseIF QS_CLEAR(Request.QueryString("op")) = "UpdatePass" Then
pold = duz(Request.Form("opass"))
pnew = duz(Request.Form("npass"))
pnew = MN_PP(pnew)
pconfirm = duz(Request.Form("cpass"))

Set mInfo = Connection.Execute("Select * FROM MEMBERS WHERE uye_id = "&Session("Uye_ID")&"")
IF pold="" OR pnew="" OR pconfirm="" Then
Response.Write "<font class=td_menu><b>"&pass_err1&"</b></font><br><br><a href=javascript:history.back();><< "&previous_page&"</a>"
ElseIF MN_PP(pold) <> mInfo("sifre") Then
Response.Write "<font class=td_menu><b>"&pass_err2&"</b></font><br><br><a href=javascript:history.back();><< "&previous_page&"</a>"
ElseIF pnew <> MN_PP(pconfirm) Then
Response.Write "<font class=td_menu><b>"&pass_err3&"</b></font><br><br><a href=javascript:history.back();><< "&previous_page&"</a>"
ELSE
Connection.Execute("UPDATE MEMBERS SET sifre = '"&pnew&"' WHERE uye_id = "&Session("Uye_ID")&"")
Response.Write "<b>" & succes_saved & "</b>"
End IF
	ElseIF QS_CLEAR(Request.QueryString("op")) = "ChangePass" Then
%>
<table border="0" cellspacing="1" cellpadding="1" width="90%" class="td_menu">
<form action="?op=UpdatePass" method="post">
<tr>
<td width="40%" align="right"><%=old_pass%> : </td>
<td width="60%"><input type="text" name="opass" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right"><%=new_pass%> : </td>
<td width="60%"><input type="text" name="npass" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right"><%=confirm_pass%> : </td>
<td width="60%"><input type="text" name="cpass" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right">&nbsp;</td>
<td width="60%"><input type="submit" value="<%=uye_guncelle%>" class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF QS_CLEAR(Request.QueryString("op")) = "RegTheme" Then
Connection.Execute("UPDATE MEMBERS SET u_theme='"&Request.QueryString("tema")&"' WHERE uye_id = "&Session("Uye_ID")&"")
Session("Theme") = Request.QueryString("tema")
Response.Redirect "your_Account.asp"
ElseIF QS_CLEAR(Request.QueryString("op")) = "Logout" Then
password=false
Session("Enter") = ""
Session("Uye_ID") = ""
Session("Member") = ""
Session("Name") = ""
Session("Email") = ""
Session("Level") = ""
Session("Signature") = ""
Session("EditPost") = ""
Session("onay") = ""
Session("Theme") = ""
Session.Abandon
Response.Cookies(""&sett("site_isim")&"")("Enter") = ""
Response.Cookies(""&sett("site_isim")&"")("gk") = ""
Response.Cookies(""&sett("site_isim")&"")("uyeid") = ""
Response.Cookies(""&sett("site_isim")&"").Expires = Now()
Response.Redirect "default.asp"

ElseIF QS_CLEAR(Request.QueryString("op")) = "Friends" Then
Set fRs = Server.CreateObject("ADODB.RecordSet")
fRsSQL = "Select * FROM FRIENDS WHERE MEMBER = "&Session("Uye_ID")&""
fRs.open frsSQL,Connection,3,3
%>
<form action="Msgbox.asp?action=fact" method="post">
<select name="fr_nm" size="10" class="form1" style="width:150">
<%
IF fRs.Eof Then
%>
<option value=""><%=no_friend%></option>
<%
ELSE
Do Until fRs.Eof
Set fr = Connection.Execute("Select * FROM MEMBERS WHERE uye_id = "&frs("FRIEND")&"")
IF fr.Eof OR fr.Bof Then
mname = "###"
ELSe
mname = fr("kul_adi")
End IF
%>
<option value="<%=mname%>"><%=mname%></option>
<%
frs.MoveNext
Loop
End If
%>
</select>
<br><br>
<input type="submit" name="fr_act" class="submit" value="<%=f_info%>">
<input type="submit" name="fr_act" class="submit" value="<%=f_msg%>">
<input type="submit" name="fr_act" class="submit" value="<%=f_delf%>">
</form>
<%
	ELSE

Set mem_l10n = Server.CreateObject("ADODB.RecordSet")
mem_l10n.open "Select * FROM NEWS WHERE ekleyen = '"&Session("Member")&"' AND onay = True ORDER BY tarih DESC",Connection,3,3
Set urMsgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND READ = False AND TO = '"&Session("Member")&"'")

If urMsgs("count") > sett("msg_limit") Then
urMsgs_COUNT = sett("msg_limit")
Else
urMsgs_COUNT = urMsgs("count")
End If
%>
<table border="0" cellspacing="2" cellpadding="0" align="center" width="100%" class="td_menu">
	<tr>
		<td valign="top">
		<%
		If uyebilgi("resmim_onay")=true Then
		resim=uyebilgi("resmim")
		Else
		resim="yok.gif"
		End if
		%>
		<a ONCLICK="window.open('upload_r.asp','upload_r','top=20,left=20,width=350,height=100,toolbar=no,scrollbars=yes');"><IMG SRC="uploads/avatar/<%=resim%>" alt="Degistimek icin tiklayin" WIDTH="120" BORDER="0" align="left"></A>Hoþgeldiniz sayin <b><%=Session("Member")%></b>
		<%
		If urMsgs_COUNT = "0" Then
		Response.Write msg_no_unreaed_msgs
		Response.Write "&nbsp;&nbsp;Resminizi guncellemek icin resmin uzerini tiklayin."
		ELSE
		%>
		<b><%=urMsgs_COUNT%></b> <%=msg_unreaded_msgs%>
		<% End If %>
		</td>
	</tr>
	<tr>
		<td valign="top">
		<hr><B>Makalelerim</B><BR>
		<%
		set sirala = server.createobject("adodb.recordset")
		SQL = "Select * from makale where gk="&uyebilgi("gk")&" order by a_title"
		sirala.open SQL,Connection,1,3
		for k=1 to sirala.RecordCount
		if sirala.eof or sirala.bof then exit for
		%>
		<A HREF="makale.asp?action=p_details&aid=<%=sirala("aid")%>"><%=sirala("a_title")%></A><br>
		<% 
		sirala.movenext : next
		sirala.close
		set sirala=nothing
		%> 
		<hr><B>Dosyalarim</B><BR>
		<%
		set sirala = server.createobject("adodb.recordset")
		SQL = "Select * from DOWNLOADS where gk="&uyebilgi("gk")&" order by p_name"
		sirala.open SQL,Connection,1,3
		for k=1 to sirala.RecordCount
		if sirala.eof or sirala.bof then exit for
		%>
		<A HREF="Programs.asp?action=p_details&pid=<%=sirala("pid")%>"><%=sirala("p_name")%></A><br>
		<% 
		sirala.movenext : next
		sirala.close
		set sirala=nothing
		%>
		<hr><B>Haberlerim</B><BR>
		<%
		set sirala = server.createobject("adodb.recordset")
		SQL = "Select * from NEWS where gk="&uyebilgi("gk")&" order by baslik"
		sirala.open SQL,Connection,1,3
		for k=1 to sirala.RecordCount
		if sirala.eof or sirala.bof then exit for
		%>
		<A HREF="news.asp?Action=Read&hid=<%=sirala("hid")%>"><%=sirala("baslik")%></A><br>
		<% 
		sirala.movenext : next
		sirala.close
		set sirala=nothing
		%>
		</td>
	</tr>
	<tr>
		<td valign="top">
		<hr><B>Forum Aboneliklerim</B><BR>
		<%
		set sirala = server.createobject("adodb.recordset")
		SQL = "Select * from f_abone where gk="&uyebilgi("gk")&""
		sirala.open SQL,Connection,1,3
		for k=1 to sirala.RecordCount
		if sirala.eof or sirala.bof then exit for
		set konu = server.createobject("adodb.recordset")
		SQL = "Select * from f_kategoriler where kategoriid="&sirala("kategoriid")&""
		konu.open SQL,Connection,1,3
		for l=1 to konu.RecordCount
		if konu.eof or konu.bof then exit for
		%>
		<A HREF="forum.asp?action=Control&id=<%=konu("kategoriid")%>"><%=konu("katad")%></A>
		<% 
		konu.movenext : next
		konu.close
		set konu=nothing
		sirala.movenext : next
		sirala.close
		set sirala=nothing
		%>
		</td>
	</tr>
</table>
<%
Response.Write "</div>"
mem_l10n.Close
Set mem_l10n = Nothing
	End IF
%>
</td>
</tr>
</table>

<%
ELSE
Response.Write WriteMsg(no_entry)
End IF
call ORTA2
call ALT
%>