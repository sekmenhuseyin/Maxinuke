<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA

action = Request.QueryString("action")
action = QS_CLEAR(action)
if action = "polls" then
call all
elseif action = "poll_details" then
call details
elseif action = "poll_process" then
call poll_process
end if

Sub all

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM POLLS ORDER BY p_date DESC"
rs.open SQL,Connection,3,3
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" bgcolor="<%=content_bg%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=top_menu6%></B></CENTER></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" class="td_menu"  bgcolor="<%=menu_bg%>">
	<tr style="font-weight:bold;">
		<td width="55%" align="left" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">Anket Sorusu</td>
	<td width="15%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%=poll_totalvotes%></td>
	<td width="30%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">Eklenme Tarihi</td>
</tr>
<%
Do Until rs.EoF
Set total_votes = Connection.Execute("SELECT SUM(a_counter) AS count FROM POLL_ALTERNATIVES WHERE pid = "&rs("poll_id")&"")
%>
<tr onMouseOver="this.style.backgroundColor='<%=forum_bg_1%>'" onMouseOut=this.style.backgroundColor="">
<td width="55%"><a href="polls.asp?action=poll_details&pid=<%=rs("poll_id")%>"><%=rs("p_question")%>
</td>
<td width="15%" align="center"><%=total_votes("count")%></td>
<td width="30%" align="center"><%=rs("p_date")%></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<!--#include file="inc/adsense.asp" -->
</td>
</tr>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub details

pollid = Request.QueryString("pid")
pollid = QS_CLEAR(pollid)
Set prs = Server.CreateObject("ADODB.RecordSet")
pSQL = "Select * FROM POLLS WHERE poll_id = "&pollid&""
prs.open pSQL,Connection,3,3

Set prs2 = Server.CreateObject("ADODB.RecordSet")
pSQL2 = "Select * FROM POLL_ALTERNATIVES WHERE pid = "&pollid&""
prs2.open pSQL2,Connection,3,3
Set prss = Connection.Execute("Select * FROM POLL_ALTERNATIVES WHERE pid = "&pollid&"")
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" bgcolor="<%=content_bg%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><B><CENTER><%=poll_details%></CENTER></B></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<center><font size="2"><%=prs("p_question")%></font></center>
<br>
<table border="0" cellspacing="5" cellpadding="5" width="100%" align="center" class="td_menu">
<%
do while not prs2.eof
Set p_tpl = Connection.Execute("SELECT SUM(a_counter) AS count FROM POLL_ALTERNATIVES WHERE pid = "&prs2("pid")&"")
If p_tpl("count") = 0 Then
percent = "0"
Else
percent = Int((prs2("a_counter") / p_tpl("count")) * 100)
End If
%>
	<tr>
		<td align="right"><b><%=prs2("alternative")%></b></td>
		<td><img src="images/ankett.gif" width="<%=percent%>" height="20"><IMG SRC="images/anket.gif"></td>
		<td><%=prs2("a_counter")%> kiþi tarafýndan seçildi, oy ortalamasý <%=percent%>%</td>
	</tr>
<%
prs2.MoveNext
Loop
%>
	<TR>
		<TD colspan="3"><%=poll_katilan%> : </b><%=p_tpl("count")%></TD>
	</TR>


</table>
<!--#include file="inc/adsense_2.asp" -->

</td>
</tr>
</table>
<%
prs.Close
set prs = Nothing
prs2.Close
set prs2 = Nothing

End Sub
Sub poll_process

pid = duz(Request.QueryString("poll_id"))
pid = QS_CLEAR(pid)
alter = duz(Request("alternative"))
IF Request.Cookies("MaxiNuke_POLL")("VOTE") = ""&pid&"" Then
Response.Write WriteMsg(poll_error0)
ELSE
Set pro = Server.CreateObject("ADODB.RecordSet")
proSQL = "Select * FROM POLL_ALTERNATIVES WHERE pid = "&pid&" AND a_id = "&alter&""
pro.open proSQL,Connection,3,3

pro("a_counter") = pro("a_counter") + 1
pro.Update

Response.Cookies("MaxiNuke_POLL")("VOTE") = ""&pid&""
Response.Cookies("MaxiNuke_POLL").Expires = Date() + 365
pro.Close
Set pro = Nothing
Response.Redirect "polls.asp?action=poll_details&pid="&pid&""
End IF
End Sub
call ORTA2
call ALT %>