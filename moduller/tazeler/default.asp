<!--#include file="../../INC/includes.asp" -->
<%
my="moduller/tazeler"
%>
<!-- ########################################################### SON 10 UYE ############################################################### -->
<table width="100%" class=td_menu>
  <tr>
    <td width="1%"><IMG SRC="<%=my%>/images/taze_uye_sol.gif"></td>
    <td width="98%"  valign="top"><p align="center"><B>En Taze Üyelerimiz</B></p>
<%
Set todaymems = Server.CreateObject("ADODB.RecordSet")
todaymems.open "Select * FROM MEMBERS WHERE uyelik_tarihi= Date()-0",Connection,3,3
For I = 1 TO 100
IF todaymems.Eof Then Exit For
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
<a href="javascript:na_open_window('win', 'kim_bu.asp?action=member_details&uid=<% response.Write ""&todaymems("uye_id")&"" %>', 300, 100, 585, 300, 0, 0, 0, 0, 0)"><% response.Write ""&todaymems("kul_adi") &""& "</a><br>"
		todaymems.MoveNext
		Next
		%></td>
    <td width="1%" align="right"><IMG SRC="<%=my%>/images/taze_uye_sag.gif"></td>
  </tr>
</table>
<!-- ########################################################### Dun uye olanlar ############################################################### -->
<hr size="1" color="<%=td_border_color%>">
<table width="100%" class=td_menu>
  <tr>
    <td width="1%"><IMG SRC="<%=my%>/images/1.gif"></td>
    <td width="98%"  valign="top"><p align="center"><B>Sitemize dün dahil olan üyelerimiz</B></p>
<%
Set todaymems = Server.CreateObject("ADODB.RecordSet")
todaymems.open "Select * FROM MEMBERS WHERE uyelik_tarihi= Date()-1",Connection,3,3
For I = 1 TO 100
IF todaymems.Eof Then Exit For
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
<a href="javascript:na_open_window('win', 'kim_bu.asp?action=member_details&uid=<% response.Write ""&todaymems("uye_id")&"" %>', 300, 100, 585, 300, 0, 0, 0, 0, 0)"><% response.Write ""&todaymems("kul_adi") &""& "</a><br>"
		todaymems.MoveNext
		Next
		%></td>
    <td width="1%" align="right"><IMG SRC="<%=my%>/images/1.gif"></td>
  </tr>
</table>
<!-- ########################################################### 2 gun once kaydolanlar ############################################################### -->
<hr size="1" color="<%=td_border_color%>">
<table width="100%" class=td_menu>
  <tr>
    <td width="1%"><IMG SRC="<%=my%>/images/4.gif"></td>
    <td width="98%"  valign="top"><p align="center"><B>Sitemize 2 gün önce Dahil olan üyelerimiz</B></p>
<%
Set todaymems = Server.CreateObject("ADODB.RecordSet")
todaymems.open "Select * FROM MEMBERS WHERE uyelik_tarihi= Date()-2",Connection,3,3
For I = 1 TO 100
IF todaymems.Eof Then Exit For
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
<a href="javascript:na_open_window('win', 'kim_bu.asp?action=member_details&uid=<% response.Write ""&todaymems("uye_id")&"" %>', 300, 100, 585, 300, 0, 0, 0, 0, 0)"><% response.Write ""&todaymems("kul_adi") &""& "</a><br>"
		todaymems.MoveNext
		Next
		%></td>
    <td width="1%" align="right"><IMG SRC="<%=my%>/images/4.gif"></td>
  </tr>
</table>

<!-- ########################################################### 3 gun once kaydolanlar ############################################################### -->
<hr size="1" color="<%=td_border_color%>">
<table width="100%" class=td_menu>
  <tr>
    <td width="1%"><IMG SRC="<%=my%>/images/6.gif"></td>
    <td width="98%"  valign="top"><p align="center"><B>Sitemize 3 gün önce Dahil olan üyelerimiz</B></p>
<%
Set todaymems = Server.CreateObject("ADODB.RecordSet")
todaymems.open "Select * FROM MEMBERS WHERE uyelik_tarihi= Date()-3",Connection,3,3
For I = 1 TO 100
IF todaymems.Eof Then Exit For
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
<a href="javascript:na_open_window('win', 'kim_bu.asp?action=member_details&uid=<% response.Write ""&todaymems("uye_id")&"" %>', 300, 100, 585, 300, 0, 0, 0, 0, 0)"><% response.Write ""&todaymems("kul_adi") &""& "</a><br>"
		todaymems.MoveNext
		Next
		%></td>
    <td width="1%" align="right"><IMG SRC="<%=my%>/images/6.gif"></td>
  </tr>
</table>

<!-- ########################################################### SON haber ############################################################### -->
<hr size="1" color="<%=td_border_color%>">
<table width="100%" class=td_menu>
	<tr>
		<td rowspan="2" width="1%"><IMG SRC="<%=my%>/images/haber_sol.gif"></td>
		<td valign="top" align="center"><B>Son eklenen 5 Haber</B></td>
		<td rowspan="2" width="1%"><IMG SRC="<%=my%>/images/haber_sag.gif"></td>
	</tr>
	<tr>
		<td width="98%">
			<%
			Set hbr5 = Connection.Execute("Select * FROM NEWS WHERE onay = True ORDER BY hid DESC")
			For last_news = 1 TO 5
			IF hbr5.eof then exit for
			%>
			<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">&nbsp;&nbsp;
			<a href="News.asp?action=Read&hid=<%=hbr5("hid")%>"> <%=duz(Left(hbr5("baslik"),90))%></a><br>
			<%
			hbr5.MoveNext
			Next%></td>
	</tr>
</table>
<!-- ########################################################### SON forum ############################################################### -->
<hr size="1" color="<%=td_border_color%>">
<table width="100%" class=td_menu>
  <tr>
    <td width="1%"><IMG SRC="<%=my%>/images/7.gif"></td>
    <td width="98%"  valign="top"><p align="center"><B>Son eklenen 5 Forum</B></p>
<%
Set l_forum = Connection.Execute("Select * FROM f_mesajlar WHERE topic = True ORDER BY mesajid DESC")
For last_forum = 1 TO 5
if l_forum.eof then exit for
Set reps = Connection.Execute("Select Count(*) AS rCount FROM f_mesajlar WHERE okunma = "&l_forum("mesajid")&" AND topic = False")
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"><a href="Forum.asp?action=Topic&id=<%=l_forum("mesajid")%>"> <%=duz(SmiLey(Left(l_forum("baslik"),70)))%></a><br>
<%
l_forum.MoveNext
Next
%>
</td>
    <td width="1%" align="right"><IMG SRC="<%=my%>/images/7.gif"></td>
  </tr>
</table>
<!-- ########################################################### SON 5 yazý ############################################################### -->
<hr size="1" color="<%=td_border_color%>">
<center><b>Son eklenen 5 yazý</b></center>
<%
Set n5 = Connection.Execute("Select * FROM makale WHERE a_approved = True ORDER BY a_date DESC")
For p = 1 To 5
if n5.eof Then exit For
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"><a href="makale.asp?action=Read&aid=<%=n5("aid")%>"> <%=duz(n5("a_title"))%></a><br>
<%
n5.MoveNext
Next
%>
<!-- ########################################################### SON download ############################################################### -->
<hr size="1" color="<%=td_border_color%>">
<center><b>Son eklenen 5 Dosya</b></center>
<%
Set n5 = Connection.Execute("Select * FROM DOWNLOADS WHERE onay = True ORDER BY pid DESC")
For p = 1 To 5
if n5.eof Then exit For
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"><a href="dosya.asp?action=p_details&pid=<%=n5("pid")%>"> <%=n5("p_name")%></a><br>
<%
n5.MoveNext
Next
%>
<!-- ########################################################### SON hikaye ############################################################### 
<hr size="1" color="<%=td_border_color%>">
<table width="100%" class=td_menu>
  <tr>
    <td width="1%"><IMG SRC="<%=my%>/images/goz.gif"></td>
    <td width="98%"  valign="top"><p align="center"><B>Son eklenen 5 Hikaye</B></p>
<%
'set conn = Server.CreateObject("ADODB.Connection")
'conn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn1W8I0R1T9E.mdb"))
'set rs = server.createobject("adodb.recordset")
'rssql = "select top 5 * from hikaye order by id DESC"
'rs.open rssql, conn, 1, 3
'do while not rs.eof
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">&nbsp;<a href="javascript:na_open_window('win', 'modules/hikaye/fikra.asp?id=<%'=rs("id")%>', 300, 100, 485, 300, 0, 0, 0, 0, 0)" target="_self"><%'=rs("fikraname")%></a><br> 
<%
'rs.movenext
'loop
'rs.close
%>
</td>
    <td width="1%" align="right"><IMG SRC="<%=my%>/images/goz.gif"></td>
  </tr>
</table>
<!-- ########################################################### SON 5 fikra ############################################################### 
<hr size="1" color="<%=td_border_color%>">
<table width="100%" class=td_menu>
  <tr>
    <td width="1%"><IMG SRC="<%=my%>/images/fikra_sol.gif"></td>
    <td width="98%"  valign="top"><p align="center"><B>Son eklenen 5 Fýkra</B></p>
<%
'set conn = Server.CreateObject("ADODB.Connection")
'conn.Open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn1W8I0R1T9E.mdb"))
'set rs = server.createobject("adodb.recordset")
'rssql = "select top 5 * from fikra WHERE onay=True  order by id DESC "
'rs.open rssql, conn, 1, 3
'do while not rs.eof
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">&nbsp;<a href="javascript:na_open_window('win', 'modules/fikra/fikra.asp?id=<%'=rs("id")%>', 300, 100, 485, 300, 0, 0, 0, 0, 0)" target="_self"><%'=rs("fikraname")%></a><br> 
<%
'rs.movenext
'loop
'rs.close
%></td>
    <td width="1%" align="right"><IMG SRC="<%=my%>/images/fikra_sag.gif"></td>
  </tr>
</table>
<!-- ########################################################### SON 5 Oyun ############################################################### 
<hr size="1" color="<%=td_border_color%>">
<center><b>Son eklenen 5 oyun</b></center>
<!-- ########################################################### SON 5 Animasyon ############################################################### 
<hr size="1" color="<%=td_border_color%>">
<center><b>Son eklenen 5 Komik Animasyon</b></center>
<!-- ########################################################### SON 5 Video ############################################################### 
<hr size="1" color="<%=td_border_color%>">
<center><b>Son eklenen 5 Video</b></center>
<!-- ########################################################### SON 5 ekart ############################################################### 
<hr size="1" color="<%=td_border_color%>">
<center><b>Son eklenen 5 E Kart</b></center>

-->