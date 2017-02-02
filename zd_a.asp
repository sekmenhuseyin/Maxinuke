<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("eylem")
If act = "" Then
call hepsi
ElseIf act = "change" Then
call change
ElseIf act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "delete" Then
call del
End If


Sub hepsi
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM GUESTBOOK ORDER BY ID DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Ziyaretçi Görüþü YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- ZÝYARETÇÝ GÖRÜÞLERÝ -=-</B></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ad Soyad</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Email</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Durum</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%Do Until rs.EoF%>
	<tr onMouseOver="this.style.backgroundColor='#3300FF'" onMouseOut=this.style.backgroundColor="">
		<td><%=rs("NAME")%></td>
		<td><%=rs("EMAIL")%></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%If rs("onay")=True then%>Aktif<%else%><font color=#339900><B>Pasif</B></font><%End if%></td>
		<td align="center"><%If rs("onay")=True then%><a href="?eylem=change&op=passive&ID=<%=rs("ID")%>">Pasifleþtir</a><%else%><a href="?eylem=change&op=active&ID=<%=rs("ID")%>">Aktifleþtir</a><%End if%> - <a href="?eylem=organize&ID=<%=rs("ID")%>">Düzenle</a> - <a href="?eylem=delete&ID=<%=rs("ID")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
End Sub

Sub change
operation = Request.QueryString("op")
id = Request.QueryString("ID")

If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If

Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM GUESTBOOK WHERE ID = "&id&""
chn.open chnSQL,Connection,1,3

chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")

chn.Close
Set chn = Nothing

End Sub

Sub organize

id = Request.QueryString("ID")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM GUESTBOOK WHERE ID = "&id&""
rs.open SQL,Connection,1,3
%>
<form action="?eylem=update&ID=<%=id%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" class="td_menu" align="center">
	<tr>
		<td>Ad ve Soyadý : </td>
		<td><input type="text" name="NAME" class="form1" value="<%=Oku(rs("NAME"))%>"></td>
		<td>E Mail Adresi : </td>
		<td><input type="text" name="EMAIL" class="form1" value="<%=Oku(rs("EMAIL"))%>"></td>
	</tr>
	<tr>
		<td colspan="4"><textarea name="YAZI" rows="10" class="form1" style="width:100%"><%=Oku(rs("YAZI"))%></textarea></td>
	</tr>
	<tr>
		<td colspan="4" align="center"><input type="submit" class="submit" value="Güncelle »" style="width:100%"></td>
	</tr>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

id = Request.QueryString("ID")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM GUESTBOOK WHERE ID = "&id&""
ent.open entSQL,Connection,1,3

NAME = Request.Form("NAME")
EMAIL = Request.Form("EMAIL")
YAZI = Request.Form("YAZI")

ent("NAME") = NAME
ent("EMAIL") = EMAIL
ent("YAZI") = YAZI
ent.Update
Response.Redirect "zd_a.asp"

ent.Close
Set ent = Nothing

End Sub
Sub del

id = Request.QueryString("ID")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM GUESTBOOK WHERE ID = "&id&""
delete.open delSQL,Connection,1,3

Set del_comments = Server.CreateObject("ADODB.RecordSet")
delCSQL = "DELETE * FROM COMMENTS WHERE nid = "&id&""
del_comments.open delCSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>