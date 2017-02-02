<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%call ORTA%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=content_bg%>">
<td align="center"><b>[</b> <a href="?action=cats">Kategoriler</a> <b>|</b> <a href="?action=Hitmakaleler">Popüler Makeleler</a> <b>|</b> <a href="?action=Newmakaleler">Yeni Eklenenler</a> <b>|</b> <a href="?action=Sendmakale">Makale Gönder</a> <b>|</b> <a href="?action=Stats">Ýstatistikler</a> <b>]</b>
<!--#include file="INC/adsense_3.asp" -->
</td>
</tr>
</table>
<br>
<%
act = Request.QueryString("action")
act = QS_CLEAR(act)
If act = "cats" Then
call cats
ElseIf act = "makaleler" Then
call makales
ElseIf act = "p_details" Then
call details
ElseIf act = "Sendmakale" Then
call makalesend
ElseIf act = "Regmakale" Then
call makalereg
ElseIf act = "Vote" Then
call pvote
ElseIF act = "Stats" Then
call stats
ElseIF act = "Hitmakaleler" Then
call hit_makales
ElseIF act = "Newmakaleler" Then
call new_makales
End If


Sub cats
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>MAKALELER</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">

<%
Set rs = Connection.Execute("Select * FROM ARTICLE_CATS ORDER BY cat_name DESC")
if rs.eof then
response.Write "<BR><BR>Eklenmiþ makale ve makale kategorisi bulunamadý.<BR><BR>Güncel makaleler için lütfen daha sonra tekrar deneyin.<BR><BR><BR>"
End if

%>
<!--#include file="inc/adsense.asp" -->
<table width="99%" align="center" cellspacing="1" class="td_menu">
<%
While Not rs.EOF

Set progs = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE cat_id = "&rs("cat_id")&" AND a_approved = True")
%>
  <tr>
    <td width="50%" height="15" bgcolor="<%=content_bg%>"><img src="images/temalar/<%=Session("SiteTheme")%>/folder.gif" align="absmiddle">&nbsp;<a href="?action=makaleler&catid=<%=rs("cat_id")%>" title="<%=rs("cat_info")%>"><b><%=rs("cat_name")%></b></a> (<%=progs("count")%>)</td>
    <%
rs.MoveNext
If Not rs.EOF Then

Set progs = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE cat_id = "&rs("cat_id")&" AND a_approved = True")
%>
    <td width="50%" height="15" bgcolor="<%=content_bg%>"><img src="images/temalar/<%=Session("SiteTheme")%>/folder.gif" align="absmiddle">&nbsp;<a href="?action=makaleler&catid=<%=rs("cat_id")%>"><b><%=rs("cat_name")%></b></a> (<%=progs("count")%>)</td>
<% End If

If Not rs.EOF Then
rs.MoveNext
End If
%>
</tr>
<% Wend %>
</table>
<br>
<%
rs.Close
Set rs = Nothing
%>
</td>
</tr>
</table>
<%
End Sub
'##############################################################################################################################
Sub makales
id = Request.QueryString("catid")
id = QS_CLEAR(id)
Set crs = Server.CreateObject("ADODB.RecordSet")
cSQL = "Select * FROM ARTICLE_CATS WHERE cat_id = "&id&""
crs.open cSQL,Connection,3,3
Set prs = Server.CreateObject("ADODB.RecordSet")
prSQL = "Select * FROM ARTICLES WHERE cat_id = "&id&" AND a_approved = True ORDER BY aid DESC"
prs.open prSQL,Connection,3,3

Page = Request.QueryString("Page")
Page = QS_CLEAR(Page)
If Page="" Then
Page = "1"
End If
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><A HREF="makaleler.asp?action=cats">Makaleler</A> / <%=crs("cat_name")%></B></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<% If prs.eof Then
Response.Write no_makale
ELSE %>
<table border="0" cellspacing="1" cellpadding="1" width="99%" bgcolor="<%=tablo_cerceve%>">
<tr>
<td bgcolor="<%=content_bg%>" class="td_menu"><!--#include file="inc/adsense.asp" -->
</td>
</tr>
<%
prs.pagesize = sett("prg_sayisi")
prs.absolutepage = Page
sayfa = prs.pagecount
For pr = 1 To prs.pagesize
If prs.eof Then Exit For
Set comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid = "&prs("aid")&" AND page = 'makaleler'")
%>
<tr>
<td bgcolor="<%=content_bg%>" class="td_menu">
<img src="images/temalar/<%=Session("SiteTheme")%>/yazi.gif" align="absmiddle">&nbsp;
<b><a href="?action=p_details&aid=<%=prs("aid")%>"><%=prs("a_title")%></a></b><br>
<font class="td_menu"><b><%=a_writer%> : </b><%=prs("a_writer")%> | <b>Toplam Okunma : </b><%=prs("hit")%> | <a href="comments.asp?action=comments&page=makaleler&id=<%=prs("aid")%>"><%=news_comments%>(<%=comments("count")%>)</a></font>
</td>
</tr>
<%
prs.MoveNext
Next
%>
</table>
<% End If %>
<br>
  <%
bp = Page-1
IF Page <> 1 Then
With Response
	.Write "<a href=""?action=makaleler&catid="&id&"&Page="&bp&""">« "&previous_page&"</a>"
	.Write "&nbsp;-&nbsp;"
End With
End IF
for y=1 to sayfa
if s=y then
response.write y
else
response.write "<a href=""?action=makaleler&catid="&id&"&Page="&y&"""><font class=""td_menu"">["&y&"]</font></a>"
end if
next
np = Page+1
IF NOT prs.Eof Then
With Response
	.Write "&nbsp;-&nbsp;"
	.Write "<a href=""?action=makaleler&catid="&id&"&Page="&np&""">"&next_page&" »</a>"
End With
End IF
%></font>
</td>
</tr>
</table>
<%
End Sub
'##############################################################################################################################































Sub hit_makales
Set l5 = Connection.Execute("Select * FROM ARTICLES WHERE a_approved = True ORDER BY hit DESC")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><CENTER>Popüler Makeleler</CENTER></B></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table border="0" cellspacing="1" cellpadding="1" width="100%" class="td_menu">
<tr><td bgcolor="<%=content_bg%>">
<%
For p = 1 To 5
if l5.eof Then exit For
Set cat = Connection.Execute("Select * FROM ARTICLE_CATS WHERE cat_id = "&l5("cat_id")&"")
%>
<font face=webdings color="<%=text%>" size=1>4</font> <b><%=l5("a_title")%></b>
<br><%=l5("a_article")%><br><b><%=makale_cat%></b> : <%=cat("cat_name")%><br><b>Toplam Okunma</b> : <%=l5("hit")%><br>
<b>[<a class="td_menu" href="?action=p_details&aid=<%=l5("aid")%>"><%=p_more%></a>]</b><hr size="1" color="<%=tablo_cerceve%>">
<%
l5.MoveNext
Next
%>
</td></tr>
</table>
<br>
<b>[</b> <a class="td_menu" href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b>
</td>
</tr>
</table>
<%
End Sub

Sub new_makales

Set n5 = Connection.Execute("Select * FROM ARTICLES WHERE a_approved = True ORDER BY a_date DESC")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><CENTER>Yeni Eklenenler</CENTER></B></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
<tr><td bgcolor="<%=content_bg%>">
<%
For p = 1 To 5
if n5.eof Then exit For
Set cat = Connection.Execute("Select * FROM ARTICLE_CATS WHERE cat_id = "&n5("cat_id")&"")
%>
<font face=webdings color="<%=text%>" size=1>4</font> <b><%=n5("a_title")%></b>
<br><%=n5("a_article")%><br><b><%=makale_cat%></b> : <%=cat("cat_name")%><br><b>Toplam Okunma</b> : <%=n5("hit")%><br>
<b>[<a class="td_menu" href="?action=p_details&aid=<%=n5("aid")%>"><%=p_more%></a>]</b><hr size="1" color="<%=tablo_cerceve%>">
<%
n5.MoveNext
Next
%>
</td></tr>
</table>
<br>
<b>[</b> <a class="td_menu" href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b>
</td>
</tr>
</table>
<%
End Sub

Sub stats

Set t_cats = Connection.Execute("Select Count(*) AS count FROM ARTICLE_CATS")
Set t_makales = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved = True")
Set makale_hits = Connection.Execute("SELECT SUM(hit) AS count FROM ARTICLES WHERE a_approved = True")
Set wait_approve = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved = False")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Ýstatistikler</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
<tr bgcolor="<%=content_bg%>" class="td_menu">
<td width="10%" align="center">
<img src="images/temalar/<%=Session("SiteTheme")%>/stats.gif">
</td>
<td width="90%">
<%=pstats_totalcats%> : <b><%=t_cats("count")%></b>
<br>
<%=pstats_totalmakales%> : <b><%=t_makales("count")%></b>
<br>
<%=pstats_totaldownload%> : <b><%=makale_hits("count")%></b>
<br>
<%=pstats_totalwa%> : <b><%=wait_approve("count")%></b>
</td>
</tr>
</table>
<br>
<b>[</b> <a class="td_menu" href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b>
</td>
</tr>
</table>
<%
End Sub



Sub details
id = Request.QueryString("aid")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ARTICLES WHERE a_approved = True AND aid = "&id&"",Connection,3,3
rs("hit") = rs("hit") + 1
rs.Update

Set comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid = "&rs("aid")&" AND page = 'makaleler'")

Set total_votes = Connection.Execute("Select SUM(point) AS count FROM ARTICLES")
IF total_votes("count") = "0" Then
makale_total_vote = "0"
ELSE
makale_total_vote = total_votes("count")
End IF
IF total_votes("count") <> "0" Then
makale_total_vote = Int((rs("point") / total_votes("count")) * 100)
End IF
set hangikat = server.createobject("adodb.recordset")
rssql = "select * from ARTICLE_CATS where cat_id = "&rs("cat_id")&""
hangikat.open rssql, Connection, 1, 3
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%"  bgcolor="<%=content_bg%>" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25">&nbsp;<B><A HREF="makaleler.asp?action=cats">Makaleler</A> / <A HREF="?action=makaleler&catid=<%=hangikat("cat_id")%>"><%=hangikat("cat_name")%></A> / <%=rs("a_title")%></B></CENTER></td>
	</tr>
<%
hangikat.close
set hangikat=nothing
%>
	<tr>
		<td valign="top" style="padding-left:10px;padding-right:10px">
<CENTER><B><%=rs("a_title")%></B></CENTER><BR>
<%Response.Write DelHTML(rs("a_article"))%>
<hr size="1" color="<%=tablo_cerceve%>">
		<CENTER><b><%=a_writer%> : </b><%=rs("a_writer")%> / - / <b>Toplam Okunma : </b><%=rs("hit")%> / <a href="comments.asp?action=comments&page=makaleler&id=<%=rs("aid")%>"><%=news_comments%>(<%=comments("count")%>)</a>
		<br><br>
		<% If Session("Enter") = "Yes" then %>
		<form action="?action=Vote&id=<%=id%>" method="post"><font class="td_menu">Makaleyi Oylayýn : </font>&nbsp;<select name="point" class="form1">
		<option value="1"><%=vote1%></option>
		<option value="2"><%=vote2%></option>
		<option value="3"><%=vote3%></option>
		<option value="4"><%=vote4%></option>
		<option value="5"><%=vote5%></option>
		</select>&nbsp;<input type="submit" value="<%=poll_button%>" class="submit">
		</form><% Else %><% End If %><font class="td_menu"><%=votes%> : <b><%=makale_total_vote%>%</b></font></CENTER>
		</td>
	</tr>
	<TR>
		<TD>Anahtar Kelimeler : <%=rs("anahtar")%><BR><!--#include file="inc/adsense_3.asp" --></TD>
	</TR>
</table>
		

<%
rs.Close
Set rs = Nothing
End Sub

Sub makalesend

Set lcats = Connection.Execute("Select * FROM ARTICLE_CATS")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Makale Gönder</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<form action="?action=Regmakale" method="post">
<table border="0" cellspacing="0" cellpadding="2" width="90%" class="td_menu">
<tr>
<td width="40%" align="right">Baþlýk : </td>
<td width="60%"><input type="text" name="a_title" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right" valign="top">Makaleniz : </td>
<td width="60%"><textarea name="a_article" rows="5" style="width=100%" class="form1"></textarea></td>
</tr>
<tr>
<td width="40%" align="right">Kategorisi : </td>
<td width="60%">
<select name="cat_id" size="1" class="form1">
<% do while not lcats.eof %>
<option value=<%=lcats("cat_id")%>><%=lcats("cat_name")%></option>
<%
lcats.MoveNext
Loop
%>
</td>
</tr>
<tr>
<td width="40%" align="right">&nbsp;</td>
<td width="60%"><input type="submit" class="submit" value="Makalemi Gönder"></td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<%
End Sub 'makalesend
Sub makalereg

a_title = duz(Request.Form("a_title"))
a_article = duz(Request.Form("a_article"))
cat_id = duz(Request.Form("cat_id"))

Set nameCheck = Connection.Execute("Select * FROM ARTICLES WHERE a_title = '"&a_title&"'")

If NOT nameCheck.Eof Then
Response.Write WriteMsg(makale_error1)
ElseIf a_title="" Then
Response.Write WriteMsg(makale_error2)
ElseIf a_article="" Then
Response.Write WriteMsg(makale_error3)
ElseIF cat_id="" Then
Response.Write WriteMsg(makale_error4)
ELSE

Set pent = Server.CreateObject("ADODB.RecordSet")
peSQL = "Select * FROM ARTICLES"
pent.open peSQL,Connection,3,3

pent.AddNew
pent("a_title") = a_title
pent("a_article") = a_article
pent("hit") = 0
pent("cat_id") = cat_id
pent("a_date") = now()
pent("a_writer") = Session("Member")
pent("point") = 0
pent("a_approved") = False
pent("align") = "left"
pent.Update
Response.Write WriteMsg(makale_succes)
END IF
END SUB 'makalereg
Sub pvote

sid = Request.QueryString("id")
sid = QS_CLEAR(sid)
votemakale = Request.Form("point")
set entvote = Server.CreateObject("ADODB.RecordSet")
evSQL = "Select * FROM ARTICLES WHERE aid = "&sid&""
entvote.open evSQL,Connection,3,3

if votemakale = "1" then
entvote("point") = entvote("point") + 1
elseif votemakale = "2" then
entvote("point") = entvote("point") + 2
elseif votemakale = "3" then
entvote("point") = entvote("point") + 3
elseif votemakale = "4" then
entvote("point") = entvote("point") + 4
elseif votemakale = "5" then
entvote("point") = entvote("point") + 5
end if
entvote.Update

Response.Redirect Request.ServerVariables("HTTP_REFERER")

End Sub 'pvote
call ORTA2
call ALT
%>