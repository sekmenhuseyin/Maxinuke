<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA 

action = Request.QueryString("action")
page = Request.QueryString("page")

IF action = "comments" then
call comments
ELSEIF action = "new_comment" then
call new_comment
ELSEIF action = "comment_register" then
call comment_reg
ELSEIF action = "delete" then
call cdel
END IF

Sub new_comment
If Session("Enter") = "Yes" Then

id = Request.QueryString("id")
id = QS_CLEAR(id)
p = Request.QueryString("page")
p = QS_CLEAR(p)
If p = "news" then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NEWS WHERE hid = "&id&"",Connection,3,3
ElseIf p = "programs" then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM DOWNLOADS WHERE cid = "&id&"",Connection,3,3
sql_conn = "Connection"
ElseIf  p = "articles" then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ARTICLES WHERE aid = "&id&"",Connection,3,3
End If
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>Yeni Yorum Ekle</B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<form action="comments.asp?page=<%=p%>&action=comment_register&id=<%=id%>" method="post">
<table border="0" cellspacing="0" cellpadding="0" width="90%" align="center" class="td_menu">
<tr><td width="40%" align="right" valign="top"><%=comment%> : </td><td width="60%"><textarea name="comment" rows="8" cols="40" class="form1"></textarea></td></tr>
<tr><td width="40%">&nbsp;</td><td width="60%" align="left"><input type="submit" value="<%=lang_button%>" class="form1"></td></tr>
</table>
</form>
</td>
</tr>
</table>
<% Else
Response.Write WriteMsg(no_entry)
End If
End Sub
Sub comment_reg
If Session("Enter") = "Yes" Then

id = Request.QueryString("id")
id = QS_CLEAR(id)
p = Request.QueryString("page")
p = QS_CLEAR(p)
com = Request.Form("comment")

Set enter = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM COMMENTS"
enter.open entSQL,Connection,3,3

IF com="" Then
Response.Write WriteMsg(empty_comment)
ELSE

enter.AddNew
enter("writer") = Session("Member")
enter("comment") = com
enter("cdate") = Now()
enter("nid") = id
enter("page") = p
enter.Update
Response.write "<div class=td_menu><b>Yorumunuz basariyla eklendi.</b></div>"
End If

End If
End Sub
Sub comments

id = Request.QueryString("id")
id = QS_CLEAR(id)
p = Request.QueryString("page")
p = QS_CLEAR(p)
If p = "news" then
Set n = Server.CreateObject("ADODB.RecordSet")
n.open "Select * FROM NEWS WHERE hid = "&id&"",Connection,3,3
ElseIf p = "programs" then
Set n = Server.CreateObject("ADODB.RecordSet")
n.open "Select * FROM DOWNLOADS WHERE pid = "&id&"",Connection,3,3
ElseIf  p = "articles" then
Set n = Server.CreateObject("ADODB.RecordSet")
n.open "Select * FROM ARTICLES WHERE aid = "&id&"",Connection,3,3
End If

Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open "Select * FROM COMMENTS WHERE nid = "&id&" AND page = '"&p&"' ORDER BY cdate DESC",Connection,3,3

If p = "news" Then
cp_title = n("baslik")
ElseIf p = "programs" Then
cp_title = n("p_name")
ElseIf p = "articles" Then
cp_title = n("a_title")
End If
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=duz(cp_title)%></B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<b><%=news_comments%></b>
<div align="left"><a href="comments.asp?page=<%=p%>&action=new_comment&id=<%=id%>"><b>» Yeni Yorum Ekle</b></a></div>
<br>
<table border="1" cellspacing="0" cellpadding="0" width="100%" bordercolor="<%=tablo_cerceve%>" bgcolor="<%=content_bg%>" class="td_menu">
<% do while not rs.eof
Set mem = Connection.Execute("Select * FROM MEMBERS Where kul_adi = '"&rs("writer")&"'")
%>
	<tr>
		<td colspan="2"><% If NOT mem.Eof Then%><a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><%end if%><%=duz(rs("writer"))%><%if NOT mem.eof Then %></a><%end if%> // <%=rs("cdate")%><% IF Session("Level")="1" OR Session("Level")="4" Then%> // <a href="?action=delete&cid=<%=rs("cid")%>">Yorumu Sil</a><% End If %><BR><%=duz(rs("comment"))%></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>

</td>
</tr>
</table>
<%
End Sub
Sub cdel
IF Session("Level")="1" OR Session("Level")="4" Then
id = Request.QueryString("cid")
id = QS_CLEAR(id)
If id<>"" then
Set dc = Server.CreateObject("ADODB.RecordSet")
dcSQL = "DELETE * FROM COMMENTS WHERE cid = "&id&""
dc.open dcSQL,Connection,3,3
End If
End If
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
call ORTA2
call ALT %>