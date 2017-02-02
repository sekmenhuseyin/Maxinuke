<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" OR Session("Level") = "4" Then
act = Request.QueryString("action")
If act = "all" Then
call all_news
ElseIf act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "new_enter" Then
call enternew
ElseIf act = "register" Then
call reg
ElseIf act = "delete" Then
call del
ElseIf act = "change" Then
call change
ElseIF act = "Cats" Then
call cats
ElseIF act = "Cat_New" Then
call cat_new
ElseIF act = "Cat_Register" Then
call cat_newreg
ElseIF act = "Cat_Edit" Then
call cat_edit
ElseIF act = "Cat_Update" Then
call cat_update
ElseIF act = "Cat_Delete" Then
call cat_delete
ElseIF act = "WaitApprove" Then
call wait_approve
ElseIF act = "AllNews" Then
call all_news
End If

Sub all_news
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM FIXED",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Haber YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- SABÝT HABERLER -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Baþlýk</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%Do Until rs.EoF%>
	<tr>
		<td align="center"><%=rs("fid")%></td>
		<td><%=rs("sf_baslik")%></td>
		<td align="center"><a href="?action=organize&fid=<%=rs("fid")%>">Düzenle</a> - <a href="?action=delete&fid=<%=rs("fid")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
%>
<%
End Sub 
Sub enternew 
%>
<form action="?action=register" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu" width="100%">
<tr>
<td align="right">Baþlýk : </td>
<td><input type="text" name="baslik" class="form1" size="60"></td>
</tr>
<tr>
<td colspan="2">
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.Value	= ""
oFCKeditor.Create "metin"
%>
</td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td align="right"><input type="submit" class="submit" size="40" value="Kaydet »"></td></tr>
</form>
</table>
<br>
<%
End Sub
Sub reg

baslik = Request("baslik")
haber = Request("metin")
If baslik="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Baþlýk Yazýnýz...</font></center>"
ElseIf haber="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Haber Metni Yazýnýz...</font></center>"
ELSE
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM FIXED",Connection,3,3

Session.LCID = 1055
ent.AddNew
ent("sf_baslik") = baslik
ent("f_metin") = haber
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=all"
End IF

End Sub
%>
<%
Sub organize

fid = Request.QueryString("fid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM FIXED WHERE fid = "&fid&""
rs.open SQL,Connection,3,3
%>
<form action="?action=update&fid=<%=fid%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" class="td_menu" align="center" width="100%">
<tr>
<td align="right">Baþlýk : </td>
<td><input type="text" name="baslik" class="form1" value="<%=Oku(rs("sf_baslik"))%>" style="width:100%"></td>
</tr>
<tr>
<td colspan="2">
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.Value	= rs("f_metin")
oFCKeditor.Create "metin"
%>
</td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td><input type="submit" class="submit" size="40" value="Güncelle »"></td>
</tr>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

fid = Request.QueryString("fid")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM FIXED WHERE fid = "&fid&""
ent.open entSQL,Connection,1,3

baslik = Request.Form("baslik")
haber = Request.Form("metin")
image = Request.Form("image")

If baslik="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Baþlýk Yazýnýz...</font></center>"
ElseIf haber="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Haber Metni Yazýnýz...</font></center>"
ELSE

ent("sf_baslik") = baslik
ent("f_metin") = haber
ent.Update
Response.Redirect "?action=all"
END IF

ent.Close
Set ent = Nothing

End Sub
Sub del

fid = Request.QueryString("fid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM FIXED WHERE fid = "&fid&""
delete.open delSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>