<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<body bgcolor="<%=background%>">
<%
Set msgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND READ = False AND TO = '"&Session("Member")&"'")

IF msgs("count") <= 0 Then
Session("New_Msg_Count") = "0"
ELSE
Session("New_Msg_Count") = msgs("count")
End IF
%>
<table border="0" cellspacing="2" cellpadding="0" align="center" width="200" class="td_menu">
<tr>
<td width="10"><img src="images/Your_Account/HaveNewMsg.gif"></td>
<td width="190">&nbsp;&nbsp;&nbsp;<%=ya_havenewmsg %></td>
</tr>
</table>
<%
msgs.Close
Set msgs = Nothing
%>