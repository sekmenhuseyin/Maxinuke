<?xml version="1.0" encoding="ISO-8859-9"?>
<%Response.Buffer = True%>
<!--#include file="inc/database.asp" -->
<!--#include file="inc/filter.asp" -->
<%
Set sett = Server.CreateObject("ADODB.RecordSet")
settSQL = "Select * FROM SETTINGS"
sett.open settSQL,Connection,1,3
Response.ContentType = "text/xml"
set dosya = server.createobject("adodb.recordset")
SQL = "Select * from ARTICLES where a_approved=true"
dosya.open SQL,Connection,1,3
%>
<rss version="2.0">
  <channel>
    <title><%=sett("site_isim")%> Makaleler</title>
    <link><%=sett("site_adres")%>/articles.asp?action=cats</link>
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
for k=1 to dosya.RecordCount
if dosya.eof or dosya.bof then exit For
set hangi_kat = server.createobject("adodb.recordset")
rssql = "select * from ARTICLE_CATS where cat_id="&dosya("cat_id")&""
hangi_kat.open rssql, Connection, 1, 3
tam=dosya("a_date")
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
      <title><![CDATA[<%=temizle(dosya("a_title"))%>]]></title>
      <link><%=sett("site_adres")%>/articles.asp?action=p_details&amp;aid=<%=dosya("aid")%></link>
      <description><%=""&temizle(Server.HTMLEncode(Left(dosya("a_article"),200)))&" . . ."%></description>
	  <author><![CDATA[<%=temizle(dosya("ekleyen"))%>]]></author>
	  <author><![CDATA[<%=temizle(sett("site_isim"))%>]]></author>
	  <category><![CDATA[<%=temizle(hangi_kat("cat_name"))%>]]></category>
  	  <pubDate><%=day(tam)%> May <%=year(tam)%> <%=right(tam,11)%> GMT</pubDate>
	  <guid><%=sett("site_adres")%>/articles.asp?action=p_details&amp;aid=<%=dosya("aid")%></guid>
	</item>
<%dosya.movenext : next%>
  </channel>
</rss>
<%
hangi_kat.close
set hangi_kat=nothing
sett.close
set sett=nothing
dosya.close
set dosya=nothing
Connection.close
set Connection=nothing
%>