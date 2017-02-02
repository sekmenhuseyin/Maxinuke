<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Or Session("Level") = "2" Or Session("Level") = "3" Or Session("Level") = "4" Or Session("Level") = "5" Then
	If Request.QueryString("eylem") = "" then
	Set kontrol = Server.CreateObject("Scripting.FileSystemObject")
		IF kontrol.FileExists(""&Server.Mappath("../db/"&date&"-"&veritabani_adi&".mdb")&"") = True Then
		response.redirect "?eylem=yedeklenmis"
		response.end
		End if
	Connection.Close
	Set Connection = Nothing
	orjinali = Server.MapPath("../db/"&veritabani_adi&".mdb")
	sikisani = Server.MapPath("../db/"&date&"-"&veritabani_adi&".mdb")
	orjinal = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& orjinali
	sikisan = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sikisani
	Set Baglanti = Server.CreateObject("JRO.JetEngine")
	Baglanti.CompactDatabase orjinal, sikisan
	Set Baglanti = Nothing
	response.redirect "?eylem=kuruldu"
	elseIf Request.QueryString("eylem") = "kuruldu" then
%>
<table width="100%" bordercolor="<%=bordercolor%>" border="1" cellspacing="3" cellpadding="3" class="td_menu">
	<tr>
		<td align="center"><B>MDB Compact (Veritabaný Sýkýþtýrma)</B></td>
	</tr>
	<tr>
		<td height="400" align="center">Veritabanlarýnýz Baþarýyla Sýkýþtýrýldý ve yedeklendi.<BR><A HREF="../db/<%=date%>-<%=veritabani_adi%>.mdb">Bilgisayarima indir</A></td>
	</tr>
</table>
<%elseIf Request.QueryString("eylem") = "yedeklenmis" then%>
<table width="100%" bordercolor="<%=bordercolor%>" border="1" cellspacing="3" cellpadding="3" class="td_menu">
	<tr>
		<td align="center"><B>MDB Compact (Veritabaný Sýkýþtýrma)</B></td>
	</tr>
	<tr>
		<td height="400" align="center">Veritabanlarýnýzi Bugun zaten yedeklemiþsiniz.</td>
	</tr>
</table>
<%
	End If
Else
response.redirect "default.asp"
End if
%>