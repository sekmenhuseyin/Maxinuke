<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Flash Oyunlar Modul Uygulamasi											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
IF QS_CLEAR(Request.QueryString("Action")) = "" Then
Set kategoriler = Server.CreateObject("ADODB.RecordSet")
kategoriler.open "Select * FROM fo_cat ORDER BY cat_name ASC",Connection,3,3
	if kategoriler.eof then
	response.Write "<CENTER><BR><BR>Veritabaninda Kayitli Flash Oyun yada Flash Oyun kategorisi bulunamadi.<BR>Lutfen daha sonra tekrar deneyin.<BR><BR><BR></CENTER>"
	End if
Do Until kategoriler.Eof
set flash_oyun_son9 = Server.CreateObJect("ADODB.RecordSet")
SQL_k = "SELECT * FROM fo WHERE cat = "&kategoriler("cat_id")&" AND onay = True ORDER BY id DESC"
flash_oyun_son9.open SQL_k,Connection,1,3
Set baks = Connection.Execute("Select SUM(oynanma) AS r_count FROM fo WHERE cat = "&kategoriler("cat_id")&"")
	IF flash_oyun_son9.RecordCount <= 0 Then
	n_baks = "0"
	ELSe
	n_baks = baks("r_count")
	End IF
baks.Close
Set baks = Nothing
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td colspan="2" align="center" height="15" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><B><%=kategoriler("cat_name")%></B></td>
	</tr>
	<tr bgcolor="<%=content_bg%>">
		<td valign="top">		
		<a href="?mi=<%=Request.QueryString("mi")%>&action=fo&id=<%=kategoriler("cat_id")%>"><img src="uploads/image/fo_cat/<%=kategoriler("cat_image")%>" border="0" align=left></a> <%=kategoriler("cat_info")%><br><hr color=<%=background%> size="1"><b>Toplam Flash Oyun : </b><%=flash_oyun_son9.RecordCount%><br><b>Toplam Oynanma : </b><%=n_baks%></td>
		<td width="60%" valign="top" rowspan="2">
		<%
		IF flash_oyun_son9.Eof Then
		Response.Write "<b>!</b> Bu Kategoride Kayýtlý Flash Oyun Yok"
		ELSE
%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
<%
for i=1 to 3
if flash_oyun_son9.eof then exit for
%>
		<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=flash_oyun_son9("id")%>"><IMG SRC="uploads/image/fo/<%=flash_oyun_son9("resim")%>" WIDTH="70" border="0"></a></TD>
<%
flash_oyun_son9.MoveNext
Next
%>
	</tr>
	<tr>
<%
for i=4 to 6
if flash_oyun_son9.eof then exit for
%>
		<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=flash_oyun_son9("id")%>"><IMG SRC="uploads/image/fo/<%=flash_oyun_son9("resim")%>" WIDTH="70" border="0"></a></TD>
<%
flash_oyun_son9.MoveNext
Next
%>
	</tr>
	<tr>
<%
for i=7 to 9
if flash_oyun_son9.eof then exit for
%>
		<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=flash_oyun_son9("id")%>"><IMG SRC="image/fo/<%=flash_oyun_son9("resim")%>" WIDTH="70" border="0"></a></TD>
<%
flash_oyun_son9.MoveNext
Next
%>
	</tr>
</table>
<%
flash_oyun_son9.Close
Set flash_oyun_son9 = Nothing
		IF I >= 10 Then
		Response.Write "<div align=""right""><a href=""?mi="&Request.QueryString("mi")&"&action=fo&id="&kategoriler("cat_id")&""">Devamý »</a></div>"
		End IF
		End IF
		%>
		</td>
	</tr>
	<tr bgcolor="<%=content_bg%>">
		<td valign="top" align="left"><!--#include file="../../inc/adsense_2.asp" --></td>
	</tr>

</table>
<br>
<%
kategoriler.MoveNext
Loop
kategoriler.Close
Set kategoriler = Nothing
Connection.Close
Set Connection = Nothing
ElseIF QS_CLEAR(Request.QueryString("action")) = "fo" Then
set fo_katici = Server.CreateObJect("ADODB.RecordSet")
SQL = "Select * FROM fo WHERE cat = "&QS_CLEAR(Request.QueryString("id"))&" and onay=true"
fo_katici.open SQL,Connection,1,3
%>
<table class="td_menu" border="0" cellspacing="5" cellpadding="0" width="100%" style="<%=TableBorder%>">
	<TR>
		<TD colspan="5"><!--#include file="../../inc/adsense.asp" --></TD>
	</TR>
<%While Not fo_katici.EOF%>
	<TR>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=fo_katici("id")%>"><IMG SRC="uploads/image/fo/<%=fo_katici("resim")%>" WIDTH="80" BORDER="0" ALT=""></a></TD>
<%
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
If Not fo_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=fo_katici("id")%>"><IMG SRC="uploads/image/fo/<%=fo_katici("resim")%>" WIDTH="80" BORDER="0" ALT=""></a></TD>
<%
End If
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
If Not fo_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=fo_katici("id")%>"><IMG SRC="uploads/image/fo/<%=fo_katici("resim")%>" WIDTH="80" BORDER="0" ALT=""></a></TD>
<%
End If
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
If Not fo_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=fo_katici("id")%>"><IMG SRC="uploads/image/fo/<%=fo_katici("resim")%>" WIDTH="80" BORDER="0" ALT=""></a></TD>
<%
End If
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
If Not fo_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=fo_katici("id")%>"><IMG SRC="uploads/image/fo/<%=fo_katici("resim")%>" WIDTH="80" BORDER="0" ALT=""></a></TD>

<%
End If
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
Wend
%>
</tr>
</table><b>Toplam Flash Oyun : </b><%=fo_katici.RecordCount%>
<%
fo_katici.Close
Set fo_katici = Nothing
Connection.Close
Set Connection = Nothing

ElseIf QS_CLEAR(Request.QueryString("Action")) = "bak" Then
fo_Voting = "oN"
id = Request.QueryString("id")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
rsSQL = "Select * FROM fo WHERE id = "&id&" AND onay = True"
rs.open rsSQL,Connection,3,3
rs("oynanma") = rs("oynanma") + 1
rs.Update
Set comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid = "&rs("id")&" AND page = 'fo'")
Set kategoriler = Connection.Execute("Select * FROM fo_cat WHERE cat_id = "&rs("cat")&"")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><A HREF="?mi=<%=Request.QueryString("mi")%>&action=fo&id=<%=kategoriler("cat_id")%>"><%=kategoriler("cat_name")%></A> / <%=rs("baslik")%></B></td>
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" align="right">-</td>
				</tr>
               <tr> 
<td align="center" colspan="2" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table border="0" cellspacing="5" cellpadding="0" align="center" width="100%" class="td_menu">
<tr>
<td valign="top" align="center">
 <OBJECT codeBase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="460" height="300" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000">
 <PARAM NAME="movie" VALUE="uploads/image/fo/<%=rs("fo_yol")%>">
 <PARAM NAME="quality" VALUE="high">
 <embed src="uploads/image/fo/<%=rs("fo_yol")%>" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="460" height="300">
 </embed>
 </OBJECT>
<BR>
<%=DelHTML(rs("metin"))%>
</td>
</tr>
</table>
<hr size="1" color="<%=tablo_cerceve%>">
<center>
Bu Flash Oyun <%=rs("oynanma")%> defa oynadý  <%=comments("count")%> defa yorumlandý.
	<hr size="1" color="<%=tablo_cerceve%>">
	<a href="comments.asp?page=fo&action=new_comment&id=<%=id%>">Yorum Yaz</a> <b>|</b>
	<a href="comments.asp?action=comments&page=fo&id=<%=id%>">Yorumlari Oku</a> <b>|</b> <b>|</b> <a href="?mi=<%=Request.QueryString("mi")%>&action=SendFriend&ST=1&id=<%=id%>"><img src="images/temalar/<%=Session("SiteTheme")%>/sendfriend.gif" border="0" align="absmiddle"> Bu Oyunu Arkadaþýna Gönder ]</a> <b>]</b>
	<hr size="1" color="<%=tablo_cerceve%>">
<!--#include file="../../inc/adsense.asp" -->
	<hr size="1" color="<%=tablo_cerceve%>">
	<%
	IF fo_Voting = "oN" AND QS_CLEAR(Request.QueryString("id")) <> "" Then
	Set nRS = Server.CreateObject("ADODB.RecordSet")
	nSQL = "Select * FROM fo WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
	nRS.open nSQL,Connection,3,3
	%>
	<%=h_averagepoint%> : <b><%=nRS("puan")%></b> - <%=h_totalvote%> : <b><%=nRS("katilimci")%></b>
	<form action="?mi=<%=Request.QueryString("mi")%>&action=Vote&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
	<input type="radio" name="n_vote" value="5" class="form1"><img src="IMAGES/site/stars-5.gif">
	<input type="radio" name="n_vote" value="4" class="form1"><img src="IMAGES/site/stars-4.gif">
	<input type="radio" name="n_vote" value="3" class="form1"><img src="IMAGES/site/stars-3.gif">
	<input type="radio" name="n_vote" value="2" class="form1"><img src="IMAGES/site/stars-2.gif">
	<input type="radio" name="n_vote" value="1" class="form1"><img src="IMAGES/site/stars-1.gif"><br><br><br>
	<input type="submit" value="<%=poll_button%>" class="submit">
	</form>
	<%
	nRS.Close
	Set nRS = Nothing
	End If
	%>
</center>
</td>
</tr>
</table>
<BR><BR>
<%
kategoriler.Close
Set kategoriler = Nothing
comments.Close
Set comments = Nothing
rs.Close
Set rs = Nothing
Connection.Close
Set Connection = Nothing

ElseIf QS_CLEAR(Request.QueryString("Action")) = "Vote" Then
IF Request.Cookies("MaxiNUKE_NV")("op") = "OK" AND Request.Cookies("MaxiNUKE_NV")("id") = ""&QS_CLEAR(Request.QueryString("id"))&"" Then
Response.Write WriteMsg("Bu Flash Oyuna önceden puan vermiþsiniz...")
ELSE
point = Request.Form("n_vote")
Set eV = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM fo WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
eV.open eSQL,Connection,3,3

eV("puan") = eV("puan") + point
eV("katilimci") = eV("katilimci") + 1
eV.Update

Response.Cookies("MaxiNUKE_NV")("op") = "OK"
Response.Cookies("MaxiNUKE_NV")("id") = ""&QS_CLEAR(Request.QueryString("id"))&""
Response.Redirect Request.ServerVariables("HTTP_REFERER")

eV.Close
Set eV = Nothing
End IF
Connection.Close
Set Connection = Nothing
ElseIF QS_CLEAR(Request.QueryString("Action")) = "SendFriend" Then
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Bu Oyunu Arkadaþýna Gönder</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<% IF QS_CLEAR(Request.QueryString("ST")) = "1" Then %>
<table border="0" cellspacing="2" cellpadding="2" width="90%" align="center" class="td_menu">
<form action="?mi=<%=Request.QueryString("mi")%>&action=SendFriend&ST=2&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<tr>
<td width="50%" align="right"><%=h_yn%> : </td>
<td width="50%"><input type="text" name="yn" size="30" class="form1" value="<%=Session("Name")%>"></td>
</tr>
<tr>
<td width="50%" align="right"><%=h_ym%> : </td>
<td width="50%"><input type="text" name="ym" size="30" class="form1" value="<%=Session("Email")%>"></td>
</tr>
<tr>
<td width="50%" align="right"><%=h_fn%> : </td>
<td width="50%"><input type="text" name="fn" size="30" class="form1"></td>
</tr>
<tr>
<td width="50%" align="right"><%=h_fm%> : </td>
<td width="50%"><input type="text" name="fm" size="30" class="form1"></td>
</tr>
<td width="50%"></td>
<td width="50%"><input type="submit" value="<%=msg_ok%>" class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF QS_CLEAR(Request.QueryString("ST")) = "2" Then
yn = duz(Request.Form("yn"))
ym = duz(Request.Form("ym"))
fn = duz(Request.Form("fn"))
fm = duz(Request.Form("fm"))
If yn="" OR ym="" OR fn="" OR fm="" Then
Response.Write advise_err
ELSE
Set nRs = Server.CreateObject("ADODB.RecordSet")
nSQL = "Select * FROM fo WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
nRs.open nSQL,Connection,3,3
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.To = ""&fm&""
objCDO.From = ""&yn&""
objCDO.Subject = ""&h_subject&""
objCDO.BodyFormat=0
objCDO.MailFormat=0

govde=govde &" <meta http-equiv=""Content-Language"" content=""tr""> "& chr(10)
govde=govde &" <meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1254""> "& chr(10)
govde=govde &" <font class=""td_menu""> "& chr(10)
govde=govde &" <b>"&yn&"</b> "&h_msg1&""& chr(10)
govde=govde &" <br><br><br> "& chr(10)
govde=govde &" "&nRs("baslik")&"<br>) "& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" "&h_msg2&" : <a href="""&sett("site_adres")&"/fo.asp?Action=bak&id="&nRs("id")&""">"&sett("site_adres")&"/fo.asp?Action=bak&id="&nRs("id")&"</a> "& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" "&h_msg3&""&sett("site_adres")&" "& chr(10)

objCDO.Body = govde
objCDO.Send
Set objCDO = Nothing
Response.Write "<font class=""td_menu"">"&h_success&"</font>"
nRs.Close
Set nRs = Nothing
End If
Connection.Close
Set Connection = Nothing
End IF
%>
</td>
</tr>
</table>
<%
End If
%>