<title>MAXI NUKE</title>
<SCRIPT LANGUAGE="JavaScript">
function SayfaYazdir()
{window.print();}
</SCRIPT>
<body onload="SayfaYazdir()" >
<!--#include file="inc/includes.asp" -->
<%
id = Request.QueryString("aid")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM makale WHERE aid = "&id&""
rs.open SQL,Connection,3,3
%>
<font face="Arial">
<center><b><%=rs("a_title")%></b><br><%=rs("a_date")%><br></center><%=DelHTML(rs("a_article"))%><br>
<center><%=h_from%> : <%=sett("site_isim")%> (<a href="<%=sett("site_adres")%>"><%=sett("site_adres")%></a>)<br><%=h_address%> : <a href="<%=sett("site_adres")%>/makale.asp?Action=Read&aid=<%=id%>"><%=sett("site_adres")%>/makale.asp?Action=Read&aid=<%=id%></a>
</center>
<%
rs.Close
Set rs = Nothing
Connection.Close
Set Connection = Nothing
%>