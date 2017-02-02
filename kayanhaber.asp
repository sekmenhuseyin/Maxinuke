<%
set tumhaberler = server.createobject("adodb.recordset")
SQL = "Select * from NEWS where onay=true order by hid desc"
tumhaberler.open SQL,Connection,1,3
%>
<table width="100%" border="0" height="16" cellspacing="0" cellpadding="0" align="center" bgcolor="<%=content_bg%>" class="td_menu" style="<%=TableBorder%>">
	<TR>
		<TD><marquee scrollamount="3" scrolldelay="70" onmouseover="this.stop();" onmouseout="this.start();">
		<%
		if tumhaberler.eof Then
		response.Write "<img src=images/site/onemli.gif height=10 width=10>&nbsp;Sisteme henüz hiç haber eklenmemiþ."
		Else
		for l=1 to tumhaberler.recordcount
		if tumhaberler.eof or tumhaberler.bof then exit For
		%>
		<a href="news.asp?Action=Read&hid=<%=tumhaberler("hid")%>">
		<img src="images/site/onemli.gif" height="10" width="10" border="0">&nbsp;<B><font size="2"><%=tumhaberler("baslik")%></a></font></B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%tumhaberler.movenext : next%>
		<%End if%></marquee></TD>
	</TR>
</TABLE>
<%
tumhaberler.close
set tumhaberler=nothing
%>