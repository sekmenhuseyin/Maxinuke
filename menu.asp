<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" Then
x = Request.QueryString("op")
	IF x = "New" Then
Set cats = Server.CreateObject("ADODB.Recordset")
cats.Open "Select * FROM MENU_CATS ORDER BY mc_title ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
<form action="?op=Enter" method="post">
	<tr>
		<td colspan="2" align="center"><B>-=- MENU LINKI EKLEME -=-</B></td>
	</tr>
<tr >
<td>Baþlýk : </td>
<td ><input type="text" name="baslik" class="form1" size="30"></td>
</tr>
<tr >
<td  >Link : </td>
<td >http://www.<input type="text" name="url" class="form1" ></td>
</tr>
<tr >
<td  >Açýlýþ Þekli : </td>
<td >
<select name="acilis" size="1" class="form1" >
<option value="_blank">Yeni Sayfa</option>
<option value="normal">Site Icinde</option>
</select>
</td>
</tr>
<tr >
<td  >Aktif/Pasif : </td>
<td  align="left"><input type="checkbox" checked name="durum"></td>
</tr>
<tr >
<td  >Kategori : </td>
<td >
<select name="kategori" size="1" class="form1" >
<% Do Until cats.EoF %>
<option value="<%=cats("mc_id")%>"><%=cats("mc_title")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr>
<td align="center" colspan="2"><input type="submit" value="Ekle »" class="submit"></td>
</tr>
</form>
</table>
<%
cats.Close
Set cats = Nothing
	ElseIF x = "Enter" Then

baslik = Replace(Request.Form("baslik"),"'","´",1,-1,1)
link = Replace(Request.Form("url"),"'","´",1,-1,1)
acilis = Request.Form("acilis")

IF Request.Form("durum") = "on" Then
durum = "True"
ELSE
durum = "False"
End IF

Set yeni = Server.CreateObject("ADODB.RecordSet")
yeni.Open "Select * FROM MENU_LINKS",Connection,3,3
yeni.AddNew
yeni("m_name") = baslik
yeni("m_url") = link
yeni("m_style") = acilis
yeni("m_status") = durum
yeni("m_cat") = Request.Form("kategori")
yeni.Update
yeni.Close
Set yeni = Nothing

Response.Redirect Request.ServerVariables("URL")

	ElseIF x = "Edit" Then
strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open "Select * FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'",Connection,3,3
IF rs("m_style") = "_blank" Then
strOpt10 = "Selected"
ELSe
strOpt11 = "Selected"
End IF
IF rs("m_status") = True Then
strOpt20 = "Checked"
End IF
Set cats = Server.CreateObject("ADODB.Recordset")
cats.Open "Select * FROM MENU_CATS ORDER BY mc_title ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
<form action="?op=Update&Name=<%=strName%>&URL=<%=strURL%>" method="post">
<tr >
<td  >Baþlýk : </td>
<td ><input type="text" name="baslik" class="form1"  value="<%=rs("m_name")%>"></td>
</tr>
<tr >
<td  >Link : </td>
<td >http://www.<input type="text" name="url" class="form1"  value="<%=rs("m_url")%>"></td>
</tr>
<tr >
<td  >Açýlýþ Þekli : </td>
<td >
<select name="acilis" size="1" class="form1" >
<option value="_blank" <%=strOpt10%>>Yeni Sayfa</option>
<option value="normal" <%=strOpt11%>>Normal Site Düzeni</option>
</select>
</td>
</tr>
<tr >
<td  >Aktif/Pasif : </td>
<td  align="left"><input type="checkbox" name="durum" <%=strOpt20%>></td>
</tr>
<tr >
<td  >Kategori : </td>
<td >
<select name="kategori" size="1" class="form1" >
<%
Do Until cats.EoF
IF cats("mc_id") = rs("m_cat") Then
opt = "Selected"
ELSE
opt = ""
End IF
%>
<option value="<%=cats("mc_id")%>" <%=opt%>><%=cats("mc_title")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr>
<td colspan="2" align="center"><input type="submit" value="Deðiþiklikleri Kaydet" class="submit"></td>
</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing
cats.Close
Set cats = Nothing
	ElseIF x = "Update" Then
strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)

baslik = Replace(Request.Form("baslik"),"'","´",1,-1,1)
link = Replace(Request.Form("url"),"'","´",1,-1,1)
acilis = Request.Form("acilis")

IF Request.Form("durum") = "on" Then
durum = "True"
ELSE
durum = "False"
End IF

Set yeni = Server.CreateObject("ADODB.RecordSet")
yeni.Open "Select * FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'",Connection,3,3

yeni("m_name") = baslik
yeni("m_url") = link
yeni("m_style") = acilis
yeni("m_status") = durum
yeni("m_cat") = Request.Form("kategori")
yeni.Update
yeni.Close
Set yeni = Nothing

Response.Redirect Request.ServerVariables("URL")
	ElseIF x = "Delete" Then
strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)
Connection.Execute("DELETE FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'")
Response.Redirect Request.ServerVariables("URL")
	ElseIF x = "Change" then

strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)
strChange = Replace(Request.QueryString("Option"),"'","´",1,-1,1)
Set upd = Server.CreateObject("ADODB.RecordSet")
upd.Open "Select * FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'",Connection,3,3
upd("m_status") = strChange
upd.Update
upd.Close
Set upd = Nothing
Response.Redirect Request.ServerVariables("URL")

	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open "Select * FROM MENU_LINKS ORDER BY m_name ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><B>-=- MENU LINKLERI -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Baþlýk</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Link</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Açýlýþ Þekli</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Durum</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
IF rs("m_status") = True Then
strDurum = "Aktif"
strYazi = "Pasifleþtir"
strOption = "False"
ELSE
strDurum = "Pasif"
strYazi = "Aktifleþtir"
strOption = "True"
End IF
IF rs("m_style") = "_blank" Then
strAcilis = "Yeni Pencere"
ELSE
strAcilis = "Normal"
End IF
strName = rs("m_name")
strURL = rs("m_url")
%>
<tr align="center">
<td align="left"><%=strName%></td>
<td ><%=strURL%></td>
<td ><%=strAcilis%></td>
<td><%=strDurum%></td>
<td ><a href="?op=Edit&Name=<%=strName%>&URL=<%=strURL%>">Düzenle</a> - <a href="?op=Delete&Name=<%=strName%>&URL=<%=strURL%>">Sil</a> - <a href="?op=Change&Name=<%=strName%>&URL=<%=strURL%>&Option=<%=strOption%>"><%=strYazi%></a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
	End IF

End IF
%>