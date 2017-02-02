<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												Videolar Modul Uygulamasi											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
IF QS_CLEAR(Request.QueryString("Action")) = "" Then
Set kategoriler = Server.CreateObject("ADODB.RecordSet")
kategoriler.open "Select * FROM video_cat ORDER BY isim ASC",Connection,3,3
	if kategoriler.eof then
	response.Write "<CENTER><BR><BR>Veritabaninda Kayitli Video yada Video kategorisi bulunamadi.<BR>Lutfen daha sonra tekrar deneyin.<BR><BR><BR></CENTER>"
	End if
Do Until kategoriler.Eof
set video_son9 = Server.CreateObJect("ADODB.RecordSet")
SQL_k = "SELECT * FROM videolar WHERE katid = "&kategoriler("katid")&" AND onay = True ORDER BY id DESC"
video_son9.open SQL_k,Connection,1,3
Set okus = Connection.Execute("Select SUM(hit) AS r_count FROM videolar WHERE katid = "&kategoriler("katid")&"")
	IF video_son9.RecordCount <= 0 Then
	n_okus = "0"
	ELSe
	n_okus = okus("r_count")
	End IF
okus.Close
Set okus = Nothing
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td align="center" height="15" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><a href="?mi=<%=Request.QueryString("mi")%>&action=video&id=<%=kategoriler("katid")%>"><B><%=kategoriler("isim")%></B></a></td>
	</tr>
<%IF video_son9.Eof Then%>
	<tr>
		<td align="center"><BR><b>!</b> Bu Kategoride Kayýtlý Video Yok<BR><BR></td>
	</tr>
<%ELSE%>
	<tr>
		<td align="center"><B><%=kategoriler("isim")%></B> kategorisi altýnda <B><%=video_son9.RecordCount%></B> adet Video bulunmakta ve bu Videolar <B><%=n_okus%></B> kez izlendiler.<hr>
			<table border="0" cellpadding="4" cellspacing="4" width="100%" class=td_menu>
				<tr>
			<%
			for i=1 to 3
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>
			<%
			for i=4 to 6
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>
			<%
			for i=7 to 9
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>
			<%
			for i=10 to 12
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>
			<%
			for i=13 to 15
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>
			<%
			for i=16 to 18
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>
			<%
			for i=19 to 21
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>			
			<%
			for i=22 to 24
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>			
			<%
			for i=25 to 27
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>			
			<%
			for i=28 to 30
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>			
			<%
			for i=31 to 33
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
				<tr>			
			<%
			for i=34 to 36
			if video_son9.eof then exit for
			%>
					<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=video_son9("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_son9("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=video_son9("buyuk")%></a></TD>
			<%
			video_son9.MoveNext
			Next
			%>
				</tr>
			</table>
<%
video_son9.Close
Set video_son9 = Nothing
IF I >= 37 Then
%>
<a href="?mi=<%=Request.QueryString("mi")%>&action=video&id=<%=kategoriler("katid")%>"><button class="submit" style="width=100%" onclick="parentElement.click();"><%=kategoriler("isim")%> kategorisi altýnda yayýnlanan tüm Videolarý göster</button></a>
<%
End IF
		End IF
		%>
		</td>
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
ElseIF QS_CLEAR(Request.QueryString("action")) = "video" Then
set fo_katici = Server.CreateObJect("ADODB.RecordSet")
SQL = "Select * FROM videolar WHERE katid = "&QS_CLEAR(Request.QueryString("id"))&" and onay=true"
fo_katici.open SQL,Connection,1,3
%>
<table class="td_menu" border="0" cellspacing="3" cellpadding="3" width="100%" style="<%=TableBorder%>">
<%While Not fo_katici.EOF%>
	<TR>
		<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=fo_katici("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=fo_katici("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=fo_katici("buyuk")%></a></TD>
<%
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
If Not fo_katici.EOF Then
%>
		<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=fo_katici("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=fo_katici("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=fo_katici("buyuk")%></a></TD>
<%
End If
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
If Not fo_katici.EOF Then
%>
		<TD width="33%" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=oku&id=<%=fo_katici("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=fo_katici("yol")%>/default.jpg"  width="120" height="90" border="0"><BR><%=fo_katici("buyuk")%></a></TD>

<%
End If
If Not fo_katici.EOF Then
fo_katici.MoveNext
End If
Wend
%>
</tr>
</table><b>Toplam Video : </b><%=fo_katici.RecordCount%>
<%
fo_katici.Close
Set fo_katici = Nothing
Connection.Close
Set Connection = Nothing

ElseIf QS_CLEAR(Request.QueryString("Action")) = "oku" Then
fo_Voting = "oN"
id = Request.QueryString("id")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
rsSQL = "Select * FROM videolar WHERE id = "&id&" AND onay = True"
rs.open rsSQL,Connection,3,3
rs("hit") = rs("hit") + 1
rs.Update
Set comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid = "&rs("id")&" AND page = 'fo'")
Set kategoriler = Connection.Execute("Select * FROM video_cat WHERE katid = "&rs("katid")&"")
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>"  class="td_menu">
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><A HREF="?mi=<%=Request.QueryString("mi")%>&action=fo&id=<%=kategoriler("katid")%>"><%=kategoriler("isim")%></A> / <%=rs("buyuk")%></B></td>
	</tr>
	<tr>
		<td align="center"><B><B><%=rs("buyuk")%></B><BR><EMBED src="http://www.youtube.com/v/<%=rs("yol")%>" width="425" height="355" type=application/x-shockwave-flash wmode="transparent"></EMBED></B><hr size="1" color="<%=tablo_cerceve%>">Bu Video <%=rs("hit")%> defa izlendi  <%=comments("count")%> defa yorumlandý.
	<hr size="1" color="<%=tablo_cerceve%>">
	<a href="comments.asp?page=fo&action=new_comment&id=<%=id%>">Yorum Yaz</a> <b>|</b>
	<a href="comments.asp?action=comments&page=fo&id=<%=id%>">Yorumlari Oku</a> <b>|</b> <b>|</b> <a href="?mi=<%=Request.QueryString("mi")%>&action=SendFriend&ST=1&id=<%=id%>"><img src="images/temalar/<%=Session("SiteTheme")%>/sendfriend.gif" border="0" align="absmiddle"> Bu Videoyu Arkadaþýna Gönder ]</a> <b>]</b>
	<hr size="1" color="<%=tablo_cerceve%>">
	<%
	IF fo_Voting = "oN" AND QS_CLEAR(Request.QueryString("id")) <> "" Then
	Set nRS = Server.CreateObject("ADODB.RecordSet")
	nSQL = "Select * FROM videolar WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
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
Response.Write WriteMsg("Bu Videoya önceden puan vermiþsiniz...")
ELSE
point = Request.Form("n_vote")
Set eV = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM videolar WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
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
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Bu Videoyu Arkadaþýna Gönder</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" class="td_menu">
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
nSQL = "Select * FROM videolar WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
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
govde=govde &" "&nRs("buyuk")&"<br>) "& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" "&h_msg2&" : <a href="""&sett("site_adres")&"/fo.asp?Action=oku&id="&nRs("id")&""">"&sett("site_adres")&"/fo.asp?Action=oku&id="&nRs("id")&"</a> "& chr(10)
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