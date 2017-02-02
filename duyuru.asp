<%
set duyuru = server.createobject("adodb.recordset")
SQL = "Select * from NOTICES ORDER BY n_id DESC"
duyuru.open SQL,Connection,1,3
if duyuru.eof then
else
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="td_menu" bgcolor="<%=content_bg%>">
	<tr>
		<td width="52"><B>Duyuru : </B></td>
		<td width="100%"><MARQUEE onmouseover="this.stop();" onmouseout="this.start();" scrollamount="3" scrolldelay="2">
		<%
		for k=1 to duyuru.RecordCount
		if duyuru.eof or duyuru.bof then exit For
		%>
		<%=duyuru("n_text")%> - 
		<%duyuru.movenext : next%>
		</td>
	</tr>
</table>
<%
End if
duyuru.close
set duyuru=nothing
%>