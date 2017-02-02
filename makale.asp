<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
IF QS_CLEAR(Request.QueryString("Action")) = "" Then
call ORTA

Set cats = Server.CreateObject("ADODB.RecordSet")
cats.open "Select * FROM makale_kat ORDER BY cat_name ASC",Connection,3,3
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu"  bgcolor="<%=content_bg%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>MAKALELER / YAZILAR</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top">
<%
if cats.eof then
response.Write "<BR><BR>Makale yada Makale kategorisi bulunamadi.<BR><BR><BR><BR>"
End if
Do Until cats.Eof
Set rs = Server.CreateObject("ADODB.RecordSet")
rsSQL = "Select * FROM makale WHERE cat = "&cats("cat_id")&" AND a_approved = True ORDER BY aid DESC"
rs.open rsSQL,Connection,3,3

Set reads = Connection.Execute("Select SUM(hit) AS r_count FROM makale WHERE cat = "&cats("cat_id")&"")
IF rs.RecordCount <= 0 Then
n_reads = "0"
ELSe
n_reads = reads("r_count")
End IF
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td colspan="2" align="center" height="15" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><B><%=cats("cat_name")%></B></td>
	</tr>
	<tr bgcolor="<%=content_bg%>">
		<td valign="top">		
		<a href="?action=makale&id=<%=cats("cat_id")%>"><img src="uploads/image/Makale_Kategorileri/<%=cats("cat_image")%>" border="0" align=left></a> <%=cats("cat_info")%><br><hr color=<%=background%> size="1"><b>Toplam Makale : </b><%=rs.RecordCount%><br><b>Toplam Okunma : </b><%=n_reads%></td>
		<td width="60%" valign="top" rowspan="2">
		<%
		IF rs.Eof Then
		Response.Write "<b>!</b> Bu Kategoride Kayýtlý Makale Yok"
		ELSE
		For I = 1 TO 14
		IF rs.Eof Then Exit For
		Response.Write "<IMG SRC=images/temalar/"&Session("SiteTheme")&"/arrow.gif>&nbsp;<a href=""?action=Read&aid="&rs("aid")&""">"&rs("a_title")&"</a> ("&rs("hit")&" kez okundu.)<br>"
		rs.MoveNext
		Next
		IF I >= 14 Then
		Response.Write "<div align=""right""><a href=""?action=makale&id="&cats("cat_id")&""">Devamý »</a></div>"
		End IF
		End IF
		%>
		</td>
	</tr>
	<tr bgcolor="<%=content_bg%>">
		<td valign="top" align="left"><!--#include file="inc/adsense_2.asp" --></td>
	</tr>

</table>
<br>
<%
cats.MoveNext
Loop
%>
</td>
</tr>
</table>
<%
call ORTA2
call ALT
ElseIF QS_CLEAR(Request.QueryString("action")) = "makale" Then
'call UST
call ORTA
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM makale WHERE cat = "&QS_CLEAR(Request.QueryString("id"))&"",Connection,3,3
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=top_menu2%></B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table class="td_menu" border="0" cellspacing="1" cellpadding="1" align="center" width="100%">
<%Do Until rs.Eof%>
	<tr>
		<td><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/yazi.gif"> <a href="?Action=Read&aid=<%=rs("aid")%>"><%=rs("a_title")%></a> (<%=rs("a_date")%>)</td>
	</tr>
<% rs.MoveNext
Loop %>
</table>
<br>
<b>Toplam Makale : </b><%=rs.RecordCount%>
</td>
</tr>
</table>
<%
call ORTA2
call ALT
ElseIf QS_CLEAR(Request.QueryString("Action")) = "Read" Then
call ORTA
makale_Voting = "oN"
id = Request.QueryString("aid")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
rsSQL = "Select * FROM makale WHERE aid = "&id&" AND a_approved = True"
rs.open rsSQL,Connection,3,3
If rs("uyelik")=True Then
	If Session("Enter") = "" then
%>
<script language="JavaScript">
alert('Bu alaný görüntüleyebilmek için üye olmanýz gerekmektedir.\n\nZaten üyeyseniz lütfen giriþ yapýn.');
</script>
<meta http-equiv="refresh" content="0;url=<%=Request.ServerVariables("HTTP_REFERER")%>">
<%
	response.end
	End if
End if
rs("hit") = rs("hit") + 1
rs.Update

Set comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid = "&rs("aid")&" AND page = 'makale'")

Set cats = Connection.Execute("Select * FROM makale_kat WHERE cat_id = "&rs("cat")&"")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><A HREF="?action=makale&id=<%=cats("cat_id")%>"><%=cats("cat_name")%></A> / <%=rs("a_title")%></B></td>
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" align="right"><a ONCLICK="window.open('Makale_yaz.asp?aid=<%=id%>','peno','top=0,left=0,width=600,height=400,toolbar=no,scrollbars=yes');"><img src="images/temalar/<%=Session("SiteTheme")%>/print.gif" border="0" align="absmiddle" alt="Bu makaleyi yazdir"></a> <a href="?Action=SendFriend&ST=1&aid=<%=id%>"><img src="images/temalar/<%=Session("SiteTheme")%>/sendfriend.gif" border="0" align="absmiddle" alt="Bu Dosyayi Arkadaþýna Gönder"></a> <%=rs("a_date")%>&nbsp;&nbsp;</td>
				</tr>
               <tr> 
<td align="center" colspan="2" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table border="0" cellspacing="5" cellpadding="0" align="center" width="100%" class="td_menu">
<tr>
<td valign="top">
<%Response.Write DelHTML(rs("a_article"))%>
</td>
</tr>
</table>
<hr size="1" color="<%=tablo_cerceve%>">
<center>
Bu makale 
<% Set nwmem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs("ekleyen")&"'") %>
    <% If NOT nwmem.Eof then %><a href="members.asp?action=member_details&uid=<%=nwmem("uye_id")%>"><%End IF %><B><%=rs("ekleyen")%></B><% IF NOT nwmem.Eof then %></a><% End If %>  tarafindan eklendi ve <%=rs("hit")%> defa okundu  <%=comments("count")%> defa yorumlandi.
	<hr size="1" color="<%=tablo_cerceve%>">
<!--#include file="inc/adsense.asp" -->
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Oylama</B></CENTER></td>
	</tr>
</table>
	<%
	IF makale_Voting = "oN" AND QS_CLEAR(Request.QueryString("aid")) <> "" Then
	Set nRS = Server.CreateObject("ADODB.RecordSet")
	nSQL = "Select * FROM makale WHERE aid = "&QS_CLEAR(Request.QueryString("aid"))&""
	nRS.open nSQL,Connection,3,3
	%>
	<%=h_averagepoint%> : <b><%=nRS("point")%></b> - <%=h_totalvote%> : <b><%=nRS("katilimci")%></b>
	<form action="?Action=Vote&aid=<%=QS_CLEAR(Request.QueryString("aid"))%>" method="post">
	<input type="radio" name="n_vote" value="5" class="form1"><img src="IMAGES/site/stars-5.gif">
	<input type="radio" name="n_vote" value="4" class="form1"><img src="IMAGES/site/stars-4.gif">
	<input type="radio" name="n_vote" value="3" class="form1"><img src="IMAGES/site/stars-3.gif">
	<input type="radio" name="n_vote" value="2" class="form1"><img src="IMAGES/site/stars-2.gif">
	<input type="radio" name="n_vote" value="1" class="form1"><img src="IMAGES/site/stars-1.gif"><br><br><br>
	<input type="submit" value="<%=poll_button%>" class="submit">
	</form>
	<%
	End if
	id = Request.QueryString("aid")
	id = QS_CLEAR(id)
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open "Select * FROM COMMENTS WHERE nid = "&id&" AND page = 'makale' ORDER BY cdate DESC",Connection,3,3
	%>
	<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
		<tr>
			<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=makale_comments%></B></CENTER></td>
		</tr>
	</table>
	<a href="comments.asp?page=makale&action=new_comment&id=<%=id%>"><b>» Yeni Yorum Ekle</b></a>
	<TABLE width="100%" class="td_menu">
	<%
	do while not rs.eof
	Set mem = Connection.Execute("Select * FROM MEMBERS Where kul_adi = '"&rs("writer")&"'")
	%>
	<TR>
		<TD><% If NOT mem.Eof Then%><a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><%end if%><%=duz(rs("writer"))%><%if NOT mem.eof Then %></a><%end if%> (<%=rs("cdate")%>)<% IF Session("Level")="1" OR Session("Level")="4" Then%> - <a href="comments.asp?action=delete&cid=<%=rs("cid")%>">Yorumu Sil</a><% End If %></TD>
	</TR>
	<TR>
		<TD><%=duz(rs("comment"))%></TD>
	</TR>
	<TR>
		<TD><hr size="1" color="<%=tablo_cerceve%>"></TD>
	</TR>
	<%
	rs.MoveNext
	Loop
	%>
	</TABLE>
</center>
</td>
</tr>
</table>
<BR><BR>
<%
call ORTA2
call ALT

ElseIf QS_CLEAR(Request.QueryString("Action")) = "Vote" Then
'call UST
call ORTA
IF Request.Cookies("MaxiNUKE_NV")("op") = "OK" AND Request.Cookies("MaxiNUKE_NV")("aid") = ""&QS_CLEAR(Request.QueryString("aid"))&"" Then
Response.Write WriteMsg(nv_errmsg)
ELSE
point = Request.Form("n_vote")
Set eV = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM makale WHERE aid = "&QS_CLEAR(Request.QueryString("aid"))&""
eV.open eSQL,Connection,3,3

eV("point") = eV("point") + point
eV("katilimci") = eV("katilimci") + 1
eV.Update

Response.Cookies("MaxiNUKE_NV")("op") = "OK"
Response.Cookies("MaxiNUKE_NV")("aid") = ""&QS_CLEAR(Request.QueryString("aid"))&""
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
call ORTA2
call ALT
ElseIF QS_CLEAR(Request.QueryString("Action")) = "SendFriend" Then
call ORTA
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Bu Dosyayi Arkadaþýna Gönder</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<% IF QS_CLEAR(Request.QueryString("ST")) = "1" Then %>
<table border="0" cellspacing="2" cellpadding="2" width="90%" align="center" class="td_menu">
<form action="?Action=SendFriend&ST=2&aid=<%=QS_CLEAR(Request.QueryString("aid"))%>" method="post">
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
nSQL = "Select * FROM makale WHERE aid = "&QS_CLEAR(Request.QueryString("aid"))&""
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
govde=govde &" "&nRs("a_title")&"<br>(Tarih : "&nRs("a_date")&") "& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" "&h_msg2&" : <a href="""&sett("site_adres")&"/makale.asp?Action=Read&aid="&nRs("aid")&""">"&sett("site_adres")&"/makale.asp?Action=Read&aid="&nRs("aid")&"</a> "& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" "&h_msg3&""&sett("site_adres")&" "& chr(10)

objCDO.Body = govde
objCDO.Send
Set objCDO = Nothing
Response.Write "<font class=""td_menu"">"&h_success&"</font>"
End IF
End IF
%>
</td>
</tr>
</table>
<%
call ORTA2
call ALT
End If
%>