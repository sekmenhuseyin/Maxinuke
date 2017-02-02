<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
If act = "all" Then
call rg
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
ElseIF act = "Allrg" Then
call all_rg
End If

Sub all_rg
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM rg ORDER BY id DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Resim YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM RESIMLER -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Baþlýk</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">bakilma</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oylanma</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Puantaj</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set cat = Connection.Execute("Select * FROM rg_cat WHERE cat_id = "&rs("cat")&"")
%>
	<tr>
		<td align="center"><%=rs("id")%></td>
		<td><%=rs("baslik")%></td>
		<td><%=cat("cat_name")%></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%=rs("bakilma")%></td>
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
Connection.Execute("DELETE FROM rg_cat WHERE cat_id = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub cat_newreg
f_1 = Request.Form("name")
f_2 = Request.Form("info")
f_3 = Request.Form("n_cat")
IF f_1="" OR f_2="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlarý Doldurunuz</div>"
ElseIF f_1<>"" AND f_2<>"" Then
IF Request.QueryString("x") = "Update" Then
Connection.Execute("UPDATE rg_cat Set cat_name = '"&f_1&"' WHERE cat_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE rg_cat Set cat_info = '"&f_2&"' WHERE cat_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE rg_cat Set cat_image = '"&f_3&"' WHERE cat_id = "&Request.QueryString("id")&"")
ELSE
Connection.Execute("INSERT INTO rg_cat (cat_name, cat_info, cat_image) VALUES ('"&f_1&"', '"&f_2&"', '"&f_3&"')")
End IF
Response.Redirect "?action=Cats"
End IF
End Sub
Sub cat_new
opt = Request.QueryString("x")
IF opt = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM rg_cat WHERE cat_id = "&Request.QueryString("id")&"",Connection,3,3
n_page = "Update"
c_name = rs("cat_name")
c_info = rs("cat_info")
c_img = Replace(rs("cat_image"),"","",1,-1,1)
c_img = Replace(c_img,"IMAGES/","",1,-1,1)
c_img = Replace(c_img,"rg_cat/","",1,-1,1)
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
<td>Kategori Açýklama : </td>
<td ><input type="text" name="info" class="form1" size="75" value="<%=c_info%>"></td>
</tr>
<tr >
<td valign="top">Kategori Resim : </td>
<td  align="right">uploads/rg_cat/<input type="text" name="n_cat" size="57" class="form1" value="<%=c_img%>"></td>
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
rs.open "Select * FROM rg_cat ORDER BY cat_name ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><B>-=- RESIM KATEGORÝLERÝ -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Resmi</td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Açýklamasý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% Do While Not rs.Eof %>
	<tr>
		<td><a href="?action=all&id=<%=rs("cat_id")%>"><IMG SRC="uploads/rg_cat/<%=rs("cat_image")%>" border="0"></a></td>
		<td><a href="?action=all&id=<%=rs("cat_id")%>"><%=rs("cat_name")%></a></td>
		<td><%=rs("cat_info")%></td>
		<td align="center"><a href="?action=Cat_New&x=Edit&id=<%=rs("cat_id")%>">Düzenle</a> - <a href="?action=Cat_Delete&id=<%=rs("cat_id")%>">Sil</a></td>
		
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

Sub rg

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM rg WHERE cat = "&Request.QueryString("id")&" AND onay = True ORDER BY id DESC"
rs.open SQL,Connection,1,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<%
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Bu Kategoride Yayýnlanan Eklenmiþ Resim YOK !</div>"
ELSE
%>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Baþlýk</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Bakilma</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oylanma</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Puantaj</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%Do While Not rs.Eof%>
	<tr>
		<td><%=rs("id")%></td>
		<td><%=rs("baslik")%></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%=rs("bakilma")%></td>
		<td align="center"><%=rs("katilimci")%></td>		
		<td align="center"><%=rs("puan")%></td>				
		<td align="center"><a href="?action=change&op=passive&id=<%=rs("id")%>">Pasifleþtir</a> - <a href="?action=organize&id=<%=rs("id")%>">Düzenle</a> - <a href="?action=delete&id=<%=rs("id")%>">Sil</a></td>
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
Set n_cats = Connection.Execute("Select * FROM rg_cat ORDER BY cat_name ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.baslik.value == "") {alert("Lütfen Baþlýk Yazýnýz... !"); return false; }
if (form.cat.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
if (form.metin.value == "") {alert("Lütfen Resim Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Resimin Baþlýðý</td>
		<td>: <input type="text" name="baslik" class="form1" size="60"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="cat" size="1" class="form1">
		<option value="0" selected>Lütfen Seçiniz</option>
		<%Do Until n_cats.Eof%>
		<option value="<%=n_cats("cat_id")%>"><%=n_cats("cat_name")%></option>
		<%
		n_cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="4">
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
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Resimi Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub
Sub reg

baslik = Request("baslik")
cat = Request("cat")
metin = Request("metin")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM rg",Connection,3,3
Session.LCID = 2048
ent.AddNew
ent("baslik") = baslik
ent("cat") = cat
ent("metin") = metin
ent("ekleyen") = Session("Member")
ent("tarih") = Now()
ent("onay") = True
ent("puan") = 0
ent("katilimci") = 0
ent("bakilma") = 0
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=all&id="&cat&""

End Sub
Sub change

operation = Request.QueryString("op")
id = Request.QueryString("id")

If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If

Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM rg WHERE id = "&id&""
chn.open chnSQL,Connection,1,3

chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")

chn.Close
Set chn = Nothing

End Sub
Sub organize

id = Request.QueryString("id")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM rg WHERE id = "&id&""
rs.open SQL,Connection,1,3

Set s_cats = Connection.Execute("Select * FROM rg_cat ORDER BY cat_name ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.baslik.value == "") {alert("Lütfen Baþlýk Yazýnýz... !"); return false; }
if (form.cat.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
if (form.metin.value == "") {alert("Lütfen Resim Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&id=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Resimin Baþlýðý</td>
		<td>: <input type="text" name="baslik" class="form1" size="60" value="<%=Oku(rs("baslik"))%>"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="cat" size="1" class="form1">
		<%
		Do Until s_cats.Eof
		IF rs("cat") = s_cats("cat_id") Then
		opt = "selected"
		ELSe
		opt = ""
		End IF
		Response.Write "<option value="""&s_cats("cat_id")&""" "&opt&">"&s_cats("cat_name")&"</option>"
		s_cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="4">
<%
Dim oFCKeditor
Set oFCKeditor = New FCKeditor
oFCKeditor.BasePath = "FCKeditor/"
oFCKeditor.Value	= rs("metin")
oFCKeditor.Create "metin"
%></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Resimi Güncelle"></td>
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
entSQL = "Select * FROM rg WHERE id = "&id&""
ent.open entSQL,Connection,1,3

baslik = Request.Form("baslik")
cat = Request.Form("cat")
metin = Request.Form("metin")
ent("baslik") = baslik
ent("cat") = cat
ent("metin") = metin
ent.Update
Response.Redirect "?action=all&id="&ent("cat")&""


ent.Close
Set ent = Nothing

End Sub
Sub del

id = Request.QueryString("id")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM rg WHERE id = "&id&""
delete.open delSQL,Connection,1,3

Set del_comments = Server.CreateObject("ADODB.RecordSet")
delCSQL = "DELETE * FROM COMMENTS WHERE nid = "&id&""
del_comments.open delCSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>