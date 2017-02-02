<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<script language="JavaScript">
<!-- // silim islemi onay uyarisi
function submitConfirm18(listForm2)
{ 
   listForm2.target="_self"; 
   listForm2.action="";
   var answer = confirm ("Silmek Istediðinize Emin misiniz ?") 
   if (answer)
       return true;
   else
       return false;
} 
//-->
</script>
<%
If Session("Level") = "1" OR Session("Level") = "4" Then
act = Request.QueryString("action")
If act = "all" Then
call news
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
ElseIF act = "AllNews" Then
call all_news
End If

Sub all_news
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NEWS ORDER BY hid DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Haber YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM HABERLER -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Baþlýk</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Okunma</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oylanma</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Puantaj</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set cat = Connection.Execute("Select * FROM NEWS_CATS WHERE cat_id = "&rs("cat")&"")
%>
	<tr onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
		<td align="center"><%=rs("hid")%></td>
		<td><a href="?action=organize&hid=<%=rs("hid")%>"><%=rs("baslik")%></a></td>
		<td><A HREF="?action=all&id=<%=cat("cat_id")%>"><%=cat("cat_name")%></A></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%=rs("okunma")%></td>
		<td align="center"><%=rs("katilimci")%></td>		
		<td align="center"><%=rs("puan")%></td>				
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&hid=<%=rs("hid")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&hid=<%=rs("hid")%>">Pasifleþtir</a> <%End if%>- <a href="?action=organize&hid=<%=rs("hid")%>">Düzenle</a> - <a onClick="return submitConfirm18(this)" href="?action=delete&hid=<%=rs("hid")%>">Sil</a></td>
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
Connection.Execute("DELETE FROM NEWS_CATS WHERE cat_id = "&Request.QueryString("id")&"")
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
Connection.Execute("UPDATE NEWS_CATS Set cat_name = '"&f_1&"' WHERE cat_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE NEWS_CATS Set cat_info = '"&f_2&"' WHERE cat_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE NEWS_CATS Set cat_image = '"&f_3&"' WHERE cat_id = "&Request.QueryString("id")&"")
ELSE
Connection.Execute("INSERT INTO NEWS_CATS (cat_name, cat_info, cat_image) VALUES ('"&f_1&"', '"&f_2&"', '"&f_3&"')")
End IF
Response.Redirect "?action=Cats"
End IF
End Sub
Sub cat_new
opt = Request.QueryString("x")
IF opt = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NEWS_CATS WHERE cat_id = "&Request.QueryString("id")&"",Connection,3,3
n_page = "Update"
c_name = rs("cat_name")
c_info = rs("cat_info")
c_img = Replace(rs("cat_image"),"","",1,-1,1)
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
<td  align="right">uploads/image/news_cats/<input type="text" name="n_cat" size="57" class="form1" value="<%=c_img%>"></td>
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
rs.open "Select * FROM NEWS_CATS ORDER BY cat_name ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><B>-=- HABER KATEGORÝLERÝ -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Resmi</td>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Açýklamasý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% Do While Not rs.Eof %>
	<tr>
		<td><a href="?action=all&id=<%=rs("cat_id")%>"><IMG SRC="uploads/image/news_cats/<%=rs("cat_image")%>" border="0"></a></td>
		<td><a href="?action=all&id=<%=rs("cat_id")%>"><%=rs("cat_name")%></a></td>
		<td><%=rs("cat_info")%></td>
		<td align="center"><a href="?action=Cat_New&x=Edit&id=<%=rs("cat_id")%>">Düzenle</a> - <a onClick="return submitConfirm18(this)" href="?action=Cat_Delete&id=<%=rs("cat_id")%>">Sil</a></td>
		
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

Sub news

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM NEWS WHERE cat = "&Request.QueryString("id")&" AND onay = True ORDER BY hid DESC"
rs.open SQL,Connection,1,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<%
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Bu Kategoride Yayýnlanan Eklenmiþ Haber YOK !</div>"
ELSE
%>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Baþlýk</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Okunma</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Oylanma</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Puantaj</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%Do While Not rs.Eof%>
	<tr onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
		<td><%=rs("hid")%></td>
		<td><a href="?action=organize&hid=<%=rs("hid")%>"><%=rs("baslik")%></a></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%=rs("okunma")%></td>
		<td align="center"><%=rs("katilimci")%></td>		
		<td align="center"><%=rs("puan")%></td>				
		<td align="center"><a href="?action=change&op=passive&hid=<%=rs("hid")%>">Pasifleþtir</a> - <a href="?action=organize&hid=<%=rs("hid")%>">Düzenle</a> - <a onClick="return submitConfirm18(this)" href="?action=delete&hid=<%=rs("hid")%>">Sil</a></td>
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
Set n_cats = Connection.Execute("Select * FROM NEWS_CATS ORDER BY cat_name ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.baslik.value == "") {alert("Lütfen Baþlýk Yazýnýz... !"); return false; }
if (form.cat.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
if (form.metin.value == "") {alert("Lütfen Haber Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Haberin Baþlýðý</td>
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
		<td>Onizleme Resim URL</td>
		<td>: <input type="text" name="image" class="form1" size="60"> (Max 400 Geniþlik)</td>
		<td>Görenler</td>
		<td>: 
		<select name="uyelik" size="1" class="form1">
		<option value=false selected>Herkes</option>
		<option value=true>Uyeler</option>
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
		%></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Haberi Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub
Sub reg
baslik = Request("baslik")
cat = Request("cat")
image = Request("image")
haber = Request("metin")
uyelik = Request("uyelik")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM NEWS",Connection,3,3
Session.LCID = 2048
ent.AddNew
ent("baslik") = baslik
ent("cat") = cat
ent("resim") = image
ent("metin") = haber
ent("ekleyen") = Session("Member")
ent("tarih") = Now()
ent("onay") = True
ent("puan") = 0
ent("katilimci") = 0
ent("okunma") = 0
ent("gk") = session("gk")
ent("uyelik") = uyelik
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=all&id="&cat&""
End Sub

Sub change
operation = Request.QueryString("op")
id = Request.QueryString("hid")

If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If

Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM NEWS WHERE hid = "&id&""
chn.open chnSQL,Connection,1,3

chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")

chn.Close
Set chn = Nothing

End Sub
Sub organize

id = Request.QueryString("hid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM NEWS WHERE hid = "&id&""
rs.open SQL,Connection,1,3

Set s_cats = Connection.Execute("Select * FROM NEWS_CATS ORDER BY cat_name ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.baslik.value == "") {alert("Lütfen Baþlýk Yazýnýz... !"); return false; }
if (form.metin.value == "") {alert("Lütfen Haber Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&hid=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Haberin Baþlýðý</td>
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
		<td>Onizleme Resim URL</td>
		<td>: <input type="text" name="image" class="form1" size="60" value="<%=rs("resim")%>"> (Max 400 Geniþlik)</td>
		<td>Görenler</td>
<%
if rs("uyelik")=false Then
uyel1="selected"
Else
uyel2="selected"
End if
%>		
		<td>: 
		<select name="uyelik" size="1" class="form1">
		<option value=false <%=uyel1%>>Herkes</option>
		<option value=true <%=uyel2%>>Uyeler</option>
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
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Guncelle"></td>
	</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

id = Request.QueryString("hid")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM NEWS WHERE hid = "&id&""
ent.open entSQL,Connection,1,3

baslik = Request.Form("baslik")
haber = Request.Form("metin")
cat = Request.Form("cat")
image = Request.Form("image")
uyelik = Request.Form("uyelik")
ent("baslik") = baslik
ent("metin") = haber
ent("cat") = cat
ent("resim") = image
ent("uyelik") = uyelik
ent.Update
Response.Redirect "?action=all&id="&ent("cat")&""


ent.Close
Set ent = Nothing

End Sub
Sub del

id = Request.QueryString("hid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM NEWS WHERE hid = "&id&""
delete.open delSQL,Connection,1,3

Set del_comments = Server.CreateObject("ADODB.RecordSet")
delCSQL = "DELETE * FROM COMMENTS WHERE nid = "&id&""
del_comments.open delCSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>