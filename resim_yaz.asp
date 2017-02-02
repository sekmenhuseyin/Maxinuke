<SCRIPT LANGUAGE="JavaScript">
function SayfaYazdir()
{window.print();}
</SCRIPT>
<body onload="SayfaYazdir()" leftmargin="0" topmargin="0">
<!--#include file="inc/includes.asp" -->
<%
id = Request.QueryString("id")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM rg WHERE id = "&id&""
rs.open SQL,Connection,3,3
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" class="td_menu">
<TR>
	<TD align="center"><b><%=rs("baslik")%></b> (<%=rs("tarih")%>)</TD>
</TR>
<TR>
	<TD align="center"><IMG SRC="images/rg/<%=rs("resim")%>"></TD>
</TR>
<TR>
	<TD align="center"><%=DelHTML(rs("metin"))%></TD>
</TR>
</TABLE>
<%
rs.Close
Set rs = Nothing
Connection.Close
Set Connection = Nothing
%>