<title>MAXINUKE</title>
<SCRIPT LANGUAGE="JavaScript">
function SayfaYazdir()
{window.print();}
</SCRIPT>
<body onload="SayfaYazdir()" >
<!--#include file="inc/includes.asp" -->
<%
id = Request.QueryString("hid")
id = QS_CLEAR(id)
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM NEWS WHERE hid = "&id&""
rs.open SQL,Connection,3,3
%>
<font face="Arial">
<center><b><%=rs("baslik")%></b><br><%=rs("tarih")%><br></center><%=DelHTML(rs("metin"))%><br>
<center><%=h_from%> : <%=sett("site_isim")%> (<a href="<%=sett("site_adres")%>"><%=sett("site_adres")%></a>)<br><%=h_address%> : <a href="<%=sett("site_adres")%>/news.asp?Action=Read&hid=<%=id%>"><%=sett("site_adres")%>/news.asp?Action=Read&hid=<%=id%></a>
</center>
<%
rs.Close
Set rs = Nothing
Connection.Close
Set Connection = Nothing
%>