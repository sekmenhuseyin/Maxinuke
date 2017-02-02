<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												Videolar Admin Uygulamasi											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
If act = "all" Then
call video
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
ElseIF act = "hepsi" Then
call hepsii
End If

Sub hepsii
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM videolar ORDER BY id DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Video YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu" >
	<tr>
		<td colspan="8" align="center"><B>-=- TUM VIDEOLAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Resmi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Videonýn Adi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Hit</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oylanma</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oy ortalamasi</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set katid = Connection.Execute("Select * FROM video_cat WHERE katid = "&rs("katid")&"")
%>
	<tr>
		<td align="center"><A HREF="http://www.youtube.com/watch?v=<%=rs("yol")%>" target="_blank"><IMG SRC="http://i.ytimg.com/vi/<%=rs("yol")%>/default.jpg" WIDTH="120" HEIGHT="90" BORDER="0" ALT=""></A></td>
		<td><%=rs("buyuk")%></td>
		<td><%=katid("isim")%></td>
		<td align="center"><%=rs("hit")%></td>
		<td align="center"><%=rs("katilimci")%></td>		
		<td align="center"><%=rs("puan")%></td>				
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&id=<%=rs("id")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&id=<%=rs("id")%>">Pasifleþtir</a> <%End if%>- <a href="?action=organize&id=<%=rs("id")%>">Düzenle</a> - <a href="?action=delete&id=<%=rs("id")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
End Sub
Sub cat_delete
Connection.Execute("DELETE FROM video_cat WHERE katid = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub cat_newreg
f_1 = Request.Form("name")
IF f_1=""Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlarý Doldurunuz</div>"
%>
<CENTER><!--#include file="inc/geri.asp" --></CENTER>
<%
ElseIF f_1<>"" Then
IF Request.QueryString("x") = "Update" Then
Connection.Execute("UPDATE video_cat Set isim = '"&f_1&"' WHERE katid = "&Request.QueryString("id")&"")
ELSE
Connection.Execute("INSERT INTO video_cat (isim) VALUES ('"&f_1&"')")
End IF
Response.Redirect "?action=Cats"
End IF
End Sub
Sub cat_new
opt = Request.QueryString("x")
IF opt = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM video_cat WHERE katid = "&Request.QueryString("id")&"",Connection,3,3
n_page = "Update"
c_name = rs("isim")
btn_submit = "Güncelle"
ELSE
c_img = "blank.gif"
btn_submit = "Ekle +"
End IF
%>
<table border="0" cellspacing="2" cellpadding="2"  align="center" class="td_menu">
<form method="post" action="?action=Cat_Register&x=<%=n_page%>&id=<%=Request.QueryString("id")%>">
<tr >
<td>Kategori Adý : </td>
<td ><input type="text" name="name" class="form1" size="75"  value="<%=c_name%>"></td>
</tr>
<tr >
<td colspan=2 align=center><input type="submit" value="<%=btn_submit%>" class="submit"></td>
</tr>
</form>
</table>
<%
End Sub
Sub cats
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM video_cat ORDER BY isim ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><B>-=- VIDEO KATEGORÝLERÝ -=-</B></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% Do While Not rs.Eof %>
	<tr>
		<td><a href="?action=all&id=<%=rs("katid")%>"><%=rs("isim")%></a></td>
		<td align="center"><a href="?action=Cat_New&x=Edit&id=<%=rs("katid")%>">Düzenle</a> - <a href="?action=Cat_Delete&id=<%=rs("katid")%>">Sil</a></td>
		
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
End Sub

Sub video

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM videolar WHERE katid = "&Request.QueryString("id")&" ORDER BY id DESC"
rs.open SQL,Connection,1,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<%
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Bu Kategoride Yayýnlanan Eklenmiþ Video YOK !</div>"
ELSE
%>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Video Adi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Hýt</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oylanma</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oy ortalamasi</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%Do While Not rs.Eof%>
	<tr>
		<td><%=rs("id")%></td>
		<td><%=rs("buyuk")%></td>
		<td align="center"><%=rs("hit")%></td>
		<td align="center"><%=rs("katilimci")%></td>		
		<td align="center"><%=rs("puan")%></td>				
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&id=<%=rs("id")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&id=<%=rs("id")%>">Pasifleþtir</a> <%End if%> - <a href="?action=organize&id=<%=rs("id")%>">Düzenle</a> - <a href="?action=delete&id=<%=rs("id")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
End if
End Sub
Sub enternew
Set n_cats = Connection.Execute("Select * FROM video_cat ORDER BY isim ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.buyuk.value == "") {alert("Lütfen Video Adini Yazýnýz... !"); return false; }
if (form.katid.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
if (form.yol.value == "") {alert("Lütfen Video yolunu Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Video Adi</td>
		<td>: <input type="text" name="buyuk" class="form1" size="60"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="katid" size="1" class="form1">
		<option value="0" selected>Lütfen Seçiniz</option>
		<%Do Until n_cats.Eof%>
		<option value="<%=n_cats("katid")%>"><%=n_cats("isim")%></option>
		<%
		n_cats.MoveNext
		Loop
		%>
		</select>
		</td>
		<td>Video Yolu</td>
		<td>: <input type="text" name="yol" class="form1" size="60"></td>
	</tr>
	<tr>
		<td align="center" colspan="6"><input type="submit" class="submit"  name="submit" style="width:100%" value="Videomý Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub
Sub reg

buyuk = Request("buyuk")
katid = Request("katid")
yol = Request("yol")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM videolar",Connection,3,3
Session.LCID = 2048
ent.AddNew
ent("buyuk") = buyuk
ent("katid") = katid
ent("yol") = yol
ent("onay") = True
ent("puan") = 0
ent("katilimci") = 0
ent("hit") = 0
ent("gonderen") = Session("Member")
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=all&id="&katid&""

End Sub


Sub organize

id = Request.QueryString("id")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM videolar WHERE id = "&id&""
rs.open SQL,Connection,1,3

Set s_cats = Connection.Execute("Select * FROM video_cat ORDER BY isim ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.buyuk.value == "") {alert("Lütfen Video Adini Yazýnýz... !"); return false; }
if (form.katid.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
if (form.yol.value == "") {alert("Lütfen Video yolunu Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&id=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Video Adi</td>
		<td>: <input type="text" name="buyuk" class="form1" size="60" value="<%=Oku(rs("buyuk"))%>"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="katid" size="1" class="form1">
		<%
		Do Until s_cats.Eof
		IF rs("katid") = s_cats("katid") Then
		opt = "selected"
		ELSe
		opt = ""
		End IF
		Response.Write "<option value="""&s_cats("katid")&""" "&opt&">"&s_cats("isim")&"</option>"
		s_cats.MoveNext
		Loop
		%>
		</select>
		</td>
		<td>Video Yolu</td>
		<td>: <input type="text" name="yol" class="form1" size="60" value="<%=Oku(rs("yol"))%>"></td>
	</tr>
	<tr>
		<td align="center" colspan="6"><input type="submit" class="submit"  name="submit" style="width:100%" value="Videomý Güncelle"></td>
	</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

id = Request.QueryString("id")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM videolar WHERE id = "&id&""
ent.open entSQL,Connection,1,3

buyuk = Request.Form("buyuk")
katid = Request.Form("katid")
yol = Request.Form("yol")

ent("buyuk") = buyuk
ent("katid") = katid
ent.Update
Response.Redirect "?action=all&id="&ent("katid")&""


ent.Close
Set ent = Nothing

End Sub


Sub change

operation = Request.QueryString("op")
If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If
Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM videolar WHERE id = "&Request.QueryString("id")&""
chn.open chnSQL,Connection,1,3
chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")
chn.Close
Set chn = Nothing
End Sub

Sub del

id = Request.QueryString("id")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM videolar WHERE id = "&id&""
delete.open delSQL,Connection,1,3

Set del_comments = Server.CreateObject("ADODB.RecordSet")
delCSQL = "DELETE * FROM COMMENTS WHERE nid = "&id&""
del_comments.open delCSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>