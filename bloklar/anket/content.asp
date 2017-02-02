<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Anket Blok Uygulamasi V 2.0												#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
Set poll = Connection.Execute("Select * FROM POLLS ORDER BY p_date DESC")
IF poll.Eof OR poll.Bof Then
Response.Write error14
ELSE
Set p_res = Connection.Execute("Select * FROM POLL_ALTERNATIVES WHERE pid = "&poll("poll_id")&"")
Set per = Connection.Execute("Select * FROM POLL_ALTERNATIVES WHERE pid = "&poll("poll_id")&"")
vote = 0
Do While Not per.EOF
vote = vote +  per("a_counter")
per.MoveNext
Loop
%>
<font class="td_menu"><%=poll("p_question")%></font>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
<%
do while not p_res.eof
Set p_tpl = Connection.Execute("SELECT SUM(a_counter) AS count FROM POLL_ALTERNATIVES WHERE pid = "&p_res("pid")&"")
	If p_res("a_counter") = "0" Then
	percent = "0"
	Else
	percent = Int((p_res("a_counter") / vote) * 100)
	End If
	IF Request.Cookies("MaxiNUKE_POLL")("VOTE") = ""&poll("poll_id")&"" Then
%>
<tr bgcolor="<%=content_bg%>"><td><b><%=p_res("alternative")%></b>&nbsp;(<%=p_res("a_counter")%> - <%=percent%>% )</td></tr>
	<% Else %>
<form action="polls.asp?action=poll_process&poll_id=<%=poll("poll_id")%>" method="post" name="form">
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="td_menu">
	<tr>
		<td width="1%"><input type="radio" name="alternative" value="<%=p_res("a_id")%>"></td>
		<td width="99%"><%=p_res("alternative")%> ( <%=p_res("a_counter")%> - <%=percent%>%)</td></tr>
<%
	End If
p_res.MoveNext
Loop
Response.Write "</table>"
	IF Request.Cookies("MaxiNUKE_POLL")("VOTE") <> ""&poll("poll_id")&"" Then
%>
<br><CENTER><input type="submit" value="<%=poll_button%>" class="submit" onClick="this.form.submit();this.disabled=true; return true;"></CENTER>
</form>
	<% End IF %>
<font class="td_menu" style="font-size:10px"><b>Bu anket </b><%=p_tpl("count")%> defa oylandi</font><br><br>
<font class="td_menu">
<b>[</b> <a href="polls.asp?action=polls"><%=poll_others%></a> <b>|</b> <a href="polls.asp?action=poll_details&pid=<%=poll("poll_id")%>"><%=poll_results%></a> <b>]</b>
</font>
<%
p_res.Close
Set p_res = Nothing
per.Close
Set per = Nothing
p_tpl.Close
Set p_tpl = Nothing
End If 
poll.Close
Set poll = Nothing
%>