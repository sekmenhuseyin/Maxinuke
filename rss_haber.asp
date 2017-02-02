<?xml version="1.0" encoding="ISO-8859-9"?>
<%Response.Buffer = True%>
<!--#include file="inc/database.asp" -->
<!--#include file="inc/filter.asp" -->
<%
Set sett = Server.CreateObject("ADODB.RecordSet")
settSQL = "Select * FROM SETTINGS"
sett.open settSQL,Connection,1,3
Response.ContentType = "text/xml"
set haber = server.createobject("adodb.recordset")
SQL = "Select * from NEWS where onay=true"
haber.open SQL,Connection,1,3
%>
<rss version="2.0">
  <channel>
    <title><%=sett("site_isim")%> Haberler</title>
    <link><%=sett("site_adres")%>/news.asp</link>
    <description><%=sett("site_slogan")%></description>
    <language>tr-tr</language>
	<pubDate><%=Date()%></pubDate>
	<lastBuildDate>%=Now()%></lastBuildDate>
	<docs><%=sett("site_adres")%></docs>
	<generator>Maxinuke</generator>
	<copyright>Her hakkı saklıdır 2008 - <%=sett("site_isim")%></copyright>
	<managingEditor><%=sett("site_mail")%></managingEditor>
    <webMaster><%=sett("site_mail")%></webMaster>
	<image>
		<url><%=sett("site_adres")%>/<%=sett("site_logo")%></url>
		<title><%=sett("site_isim")%></title>
		<link><%=sett("site_adres")%>/</link>
		<width>310</width>
		<height>60</height>
	</image>
<%
for k=1 to haber.RecordCount
if haber.eof or haber.bof then exit For
tam=haber("tarih")
ay=month(tam)
If ay = 1 Then
ayy = "Jan"
elseIf ay = 2 Then
ayy = "Feb"
elseIf ay = 3 Then
ayy = "Mar"
elseIf ay = 4 Then
ayy = "Apr"
elseIf ay = "5" Then
ayy = "May"
elseIf ay = 6 Then
ayy = "Jun"
elseIf ay = 7 Then
ayy = "Jul"
elseIf ay = 8 Then
ayy = "Aug"
elseIf ay = 9 Then
ayy = "Sep"
elseIf ay = 10 Then
ayy = "Oct"
elseIf ay = 11 Then
ayy = "Nov"
elseIf ay = 12 Then
ayy = "Dec"
End If
%>
    <item>
      <title><![CDATA[<%=temizle(haber("baslik"))%>]]></title>
      <link><%=sett("site_adres")%>/news.asp?Action=Read&amp;hid=<%=haber("hid")%></link>
      <description><%=""&temizle(Server.HTMLEncode(Left(haber("metin"),200)))&" . . ."%></description>
	  <author><![CDATA[<%=temizle(haber("ekleyen"))%>]]></author>
  	  <pubDate><%=day(tam)%> May <%=year(tam)%> <%=right(tam,11)%> GMT</pubDate>
	  <guid><%=sett("site_adres")%>/news.asp?Action=Read&amp;hid=<%=haber("hid")%></guid>
	</item>
<%haber.movenext : next%>
  </channel>
</rss>
<%
sett.close
set sett=nothing
haber.close
set haber=nothing
Connection.close
set Connection=nothing
%>