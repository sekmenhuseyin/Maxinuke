<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Resim Galerisi Modul Uygulamasi V 1.1										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
IF QS_CLEAR(Request.QueryString("Action")) = "" Then
Set cats = Server.CreateObject("ADODB.RecordSet")
cats.open "Select * FROM rg_cat ORDER BY cat_name ASC",Connection,3,3
	if cats.eof then
	response.Write "<CENTER><BR><BR>Veritabaninda Kayitli Resim yada Resim kategorisi bulunamadi.<BR>Lutfen daha sonra tekrar deneyin.<BR><BR><BR></CENTER>"
	End if
Do Until cats.Eof
set rgson12 = Server.CreateObJect("ADODB.RecordSet")
SQL_k = "SELECT * FROM rg WHERE cat = "&cats("cat_id")&" AND onay = True ORDER BY id DESC"
rgson12.open SQL_k,Connection,1,3
Set baks = Connection.Execute("Select SUM(bakilma) AS r_count FROM rg WHERE cat = "&cats("cat_id")&"")
	IF rgson12.RecordCount <= 0 Then
	n_baks = "0"
	ELSe
	n_baks = baks("r_count")
	End IF
baks.Close
Set baks = Nothing
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td colspan="2" align="center" height="15" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><B><%=cats("cat_name")%></B></td>
	</tr>
	<tr bgcolor="<%=content_bg%>">
		<td valign="top">		
		<a href="?mi=<%=Request.QueryString("mi")%>&action=rg&id=<%=cats("cat_id")%>"><img src="uploads/image/rg_cat/<%=cats("cat_image")%>" border="0" align=left></a> <%=cats("cat_info")%><br><hr color=<%=background%> size="1"><b>Toplam Resim : </b><%=rgson12.RecordCount%><br><b>Toplam Görüntülenme : </b><%=n_baks%></td>
		<td width="60%" valign="top" rowspan="2">
		<%
		IF rgson12.Eof Then
		Response.Write "<b>!</b> Bu Kategoride Kayýtlý Resim Yok"
		ELSE
%>
<table border="0" cellpadding="0" cellspacing="5" width="100%">
	<tr>
<%
for i=1 to 3
if rgson12.eof then exit for
%>
		<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rgson12("id")%>&cat=<%=rgson12("cat")%>"><%=rgson12("baslik")%></a></TD>
<%
rgson12.MoveNext
Next
%>
	</tr>
	<tr>
<%
for i=4 to 6
if rgson12.eof then exit for
%>
		<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rgson12("id")%>&cat=<%=rgson12("cat")%>"><%=rgson12("baslik")%></a></TD>
<%
rgson12.MoveNext
Next
%>
	</tr>
	<tr>
<%
for i=7 to 9
if rgson12.eof then exit for
%>
		<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rgson12("id")%>&cat=<%=rgson12("cat")%>"><%=rgson12("baslik")%></a></TD>
<%
rgson12.MoveNext
Next
%>
	</tr>
	<tr>
<%
for i=10 to 12
if rgson12.eof then exit for
%>
		<TD align="center" style="<%=TableBorder%>"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rgson12("id")%>&cat=<%=rgson12("cat")%>"><%=rgson12("baslik")%></a></TD>

<%
rgson12.MoveNext
Next
%>
	</tr>
</table>
<%
rgson12.Close
Set rgson12 = Nothing
		IF I >= 10 Then
		Response.Write "<div align=""right""><a href=""?mi="&Request.QueryString("mi")&"&action=rg&id="&cats("cat_id")&""">Devamý »</a></div>"
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
cats.MoveNext
Loop
cats.Close
Set cats = Nothing
Connection.Close
Set Connection = Nothing
ElseIF QS_CLEAR(Request.QueryString("action")) = "rg" Then
set rg_katici = Server.CreateObJect("ADODB.RecordSet")
SQL = "Select * FROM rg WHERE cat = "&QS_CLEAR(Request.QueryString("id"))&" and onay=true"
rg_katici.open SQL,Connection,1,3
%>
<table class="td_menu" border="0" cellspacing="5" cellpadding="0" width="100%" style="<%=TableBorder%>">
	<TR>
		<TD colspan="5"><!--#include file="../../inc/adsense.asp" --></TD>
	</TR>
<%While Not rg_katici.EOF%>
	<TR>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rg_katici("id")%>&cat=<%=rg_katici("cat")%>"><%=rg_katici("baslik")%></a></TD>
<%
If Not rg_katici.EOF Then
rg_katici.MoveNext
End If
If Not rg_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rg_katici("id")%>&cat=<%=rg_katici("cat")%>"><%=rg_katici("baslik")%></a></TD>
<%
End If
If Not rg_katici.EOF Then
rg_katici.MoveNext
End If
If Not rg_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rg_katici("id")%>&cat=<%=rg_katici("cat")%>"><%=rg_katici("baslik")%></a></TD>
<%
End If
If Not rg_katici.EOF Then
rg_katici.MoveNext
End If
If Not rg_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rg_katici("id")%>&cat=<%=rg_katici("cat")%>"><%=rg_katici("baslik")%></a></TD>
<%
End If
If Not rg_katici.EOF Then
rg_katici.MoveNext
End If
If Not rg_katici.EOF Then
%>
		<TD style="<%=TableBorder%>" align="center"><a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=rg_katici("id")%>&cat=<%=rg_katici("cat")%>"><%=rg_katici("baslik")%></a></TD>

<%
End If
If Not rg_katici.EOF Then
rg_katici.MoveNext
End If
Wend
%>
</tr>
</table><b>Toplam Resim : </b><%=rg_katici.RecordCount%>
<%
rg_katici.Close
Set rg_katici = Nothing
Connection.Close
Set Connection = Nothing

ElseIf QS_CLEAR(Request.QueryString("Action")) = "bak" Then
rg_Voting = "oN"
id = Request.QueryString("id")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
rsSQL = "Select * FROM rg WHERE id = "&id&" AND onay = True"
rs.open rsSQL,Connection,3,3
rs("bakilma") = rs("bakilma") + 1
rs.Update
Set comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid = "&rs("id")&" AND page = 'rg'")
Set cats = Connection.Execute("Select * FROM rg_cat WHERE cat_id = "&rs("cat")&"")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><A HREF="?mi=<%=Request.QueryString("mi")%>&action=rg&id=<%=cats("cat_id")%>"><%=cats("cat_name")%></A> / <%=rs("baslik")%></B></td>
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" align="right"><%=rs("tarih")%>&nbsp;&nbsp;</td>
				</tr>
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25" colspan="2">
					<%
					id = Request.QueryString("id")
					geriidi = id-1
					ileriidi = id+1					
					
					
					
					
					
					
					
					
					
					
					
					cat = Request.QueryString("cat")
					Eylem = Request.QueryString("eylem")

					
	
					

					Set gerisi = Connection.Execute("Select id From rg where cat="&cat&"")

					If gerisi("id") = CInt(id) Then Sira = i
					sonID = gerisi("id")

					gerisi.Close
					Set gerisi = Nothing




					
					
					
					Set ilerigeri = Connection.Execute("Select id From rg where cat="&cat&"")
					ilkID = ilerigeri("id")
					i = 0
					Do While not ilerigeri.Eof

					If ilerigeri("id") = CInt(id) Then Sira = i
					sonID = ilerigeri("id")
					i = i + 1
					ilerigeri.MoveNext
					Loop
					ilerigeri.Close
					Set ilerigeri = Nothing

					Set ilerigeri = Connection.Execute("Select * From rg where cat="&Request.QueryString("cat")&"")
					If Eylem = "sonraki" Then
					ilerigeri.Move(Sira+1)
					ElseIf Eylem = "onceki" Then
					ilerigeri.Move(Sira-1)
					End If
					%>
					<TABLE width="100%" border="0" class="td_menu">
					<TR>
						<TD align="center">


<%=geriidi%>
<%=ileriidi%>



						<% If ilerigeri("id") <> CInt(ilkID) Then %>
						<a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=ilerigeri("id")%>&cat=<%=ilerigeri("cat")%>&eylem=onceki"><<<< Önceki <<<<</a>
						<% End If %>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<% If ilerigeri("id") <> CInt(sonID) Then %>
						<a href="?mi=<%=Request.QueryString("mi")%>&action=bak&id=<%=ilerigeri("id")%>&cat=<%=ilerigeri("cat")%>&eylem=sonraki">>>> Sonraki >>></a>
						<% End If %>
						</TD>
					</TR>
					</TABLE>
					<% 
					ilerigeri.Close
					Set ilerigeri = Nothing
					%>
				  </td>
				</tr>
			   <tr> 
<td align="center" colspan="2" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table border="0" cellspacing="5" cellpadding="0" align="center" width="100%" class="td_menu">
<tr>
<td valign="top" align="center"><%=rs("metin")%></td>
</tr>
</table>
<hr size="1" color="<%=tablo_cerceve%>">
<center>
Bu Resim 
<% Set nwmem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs("ekleyen")&"'") %>
    <% If NOT nwmem.Eof then %><a href="members.asp?action=member_details&uid=<%=nwmem("uye_id")%>"><%End IF %><B><%=rs("ekleyen")%></B><% IF NOT nwmem.Eof then %></a><% End If %>  tarafindan eklendi ve <%=rs("bakilma")%> defa görüntülendi  <%=comments("count")%> defa yorumlandi.
	<hr size="1" color="<%=tablo_cerceve%>">
	<a href="comments.asp?page=rg&action=new_comment&id=<%=id%>">Yorum Yaz</a> <b>|</b>
	<a href="comments.asp?action=comments&page=rg&id=<%=id%>">Yorumlari Oku</a> <b>|</b> <a ONCLICK="window.open('resim_yaz.asp?id=<%=id%>','peno','top=0,left=0,width=600,height=400,toolbar=no,scrollbars=yes');"><img src="images/temalar/<%=Session("SiteTheme")%>/print.gif" border="0" align="absmiddle"> Bu Resmi Yazdýr </a> <b>|</b> <a href="?mi=<%=Request.QueryString("mi")%>&action=SendFriend&ST=1&id=<%=id%>"><img src="images/temalar/<%=Session("SiteTheme")%>/sendfriend.gif" border="0" align="absmiddle"> Bu Resmi Arkadaþýna Gönder ]</a> <b>]</b>
	<hr size="1" color="<%=tablo_cerceve%>">
<!--#include file="../../inc/adsense.asp" -->
	<hr size="1" color="<%=tablo_cerceve%>">
	<%
	IF rg_Voting = "oN" AND QS_CLEAR(Request.QueryString("id")) <> "" Then
	Set nRS = Server.CreateObject("ADODB.RecordSet")
	nSQL = "Select * FROM rg WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
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
nwmem.Close
Set nwmem = Nothing
cats.Close
Set cats = Nothing
comments.Close
Set comments = Nothing
rs.Close
Set rs = Nothing
Connection.Close
Set Connection = Nothing

ElseIf QS_CLEAR(Request.QueryString("Action")) = "Vote" Then
IF Request.Cookies("MaxiNUKE_NV")("op") = "OK" AND Request.Cookies("MaxiNUKE_NV")("id") = ""&QS_CLEAR(Request.QueryString("id"))&"" Then
Response.Write WriteMsg("Bu Resime önceden puan vermiþsiniz...")
ELSE
point = Request.Form("n_vote")
Set eV = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM rg WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
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
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Bu Resmi Arkadaþýna Gönder</B></CENTER></td>
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
nSQL = "Select * FROM rg WHERE id = "&QS_CLEAR(Request.QueryString("id"))&""
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
govde=govde &" "&nRs("baslik")&"<br>(Tarih : "&nRs("tarih")&") "& chr(10)
govde=govde &" <br><br> "& chr(10)
govde=govde &" "&h_msg2&" : <a href="""&sett("site_adres")&"/rg.asp?Action=bak&id="&nRs("id")&""">"&sett("site_adres")&"/rg.asp?Action=bak&id="&nRs("id")&"</a> "& chr(10)
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