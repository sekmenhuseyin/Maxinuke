<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
If act = "yenisayfa" Then
call yenisayfaa
ElseIf act = "kayit" Then
call kayitt
ElseIf act = "hepsi" Then
call hepsii
ElseIf act = "change" Then
call change
ElseIf act = "IS" Then
call ImageSelect
ElseIf act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "IU" Then
call ImageUpload
ElseIf act = "delete" Then
call del

End If
'################################################################################################################################
Sub yenisayfaa
Set n_p_cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.p_title.value == "") {alert("Lütfen Baþlýk Yazýnýz... !"); return false; }
if (form.p_cat.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
if (form.p_content.value == "") {alert("Lütfen Sayfa Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=kayit" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Sayfanin Baþlýðý</td>
		<td>: <input type="text" name="p_title" class="form1" size="60"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="p_cat" size="1" class="form1">
		<option value="0" selected>Lütfen Seçiniz</option>
		<%Do Until n_p_cats.Eof%>
		<option value="<%=n_p_cats("mc_id")%>"><%=n_p_cats("mc_title")%></option>
		<%
		n_p_cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="3">-</td>
		<td>Kimler Gorebilir
		<select name="mems" size="1" class="form1">
		<option value=true>Sadece Uyeler</option>
		<option value=false>Herkes</option>
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
oFCKeditor.Create "p_content"
%></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Sayfani Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub
'################################################################################################################################
Sub kayitt
p_title = Request("p_title")
p_cat = Request("p_cat")
mems = Request("mems")
p_content = Request("p_content")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM PAGES",Connection,3,3
Session.LCID = 2048
ent.AddNew
ent("p_title") = p_title
ent("p_cat") = p_cat
ent("mems") = mems
ent("p_content") = p_content
ent("onay") = True
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=hepsi"
End Sub
'################################################################################################################################
Sub hepsii
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM PAGES ORDER BY p_id DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Sayfa YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM SAYFALAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Baþlýk</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kimler Gorur</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set p_cat = Connection.Execute("Select * FROM MENU_CATS WHERE mc_id = "&rs("p_cat")&"")
%>
	<tr>
		<td align="center"><%=rs("p_id")%></td>
		<td><%=rs("p_title")%></td>
		<td><%=p_cat("mc_title")%></td>
		<td align="center"><%If rs("mems") = false then %>Herkes <%else%><font color="#FF0000"><B>Sadece Uyeler</B> <%End if%></td>
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&p_id=<%=rs("p_id")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&p_id=<%=rs("p_id")%>">Pasifleþtir</a> <%End if%>- <a href="?action=organize&p_id=<%=rs("p_id")%>">Düzenle</a> - <a href="?action=delete&p_id=<%=rs("p_id")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
End Sub
'################################################################################################################################
Sub change
operation = Request.QueryString("op")
id = Request.QueryString("p_id")
If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If
Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM PAGES WHERE p_id = "&id&""
chn.open chnSQL,Connection,1,3
chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")
chn.Close
Set chn = Nothing
End Sub
'################################################################################################################################
Sub organize
id = Request.QueryString("p_id")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM PAGES WHERE p_id = "&id&""
rs.open SQL,Connection,1,3
Set s_p_cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.p_title.value == "") {alert("Lütfen Baþlýk Yazýnýz... !"); return false; }
if (form.p_content.value == "") {alert("Lütfen sayfa Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&p_id=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Sayfain Baþlýðý</td>
		<td>: <input type="text" name="p_title" class="form1" size="60" value="<%=Oku(rs("p_title"))%>"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="p_cat" size="1" class="form1">
		<%
		Do Until s_p_cats.Eof
		IF rs("p_cat") = s_p_cats("mc_id") Then
		opt = "selected"
		ELSe
		opt = ""
		End IF
		Response.Write "<option value="""&s_p_cats("mc_id")&""" "&opt&">"&s_p_cats("mc_title")&"</option>"
		s_p_cats.MoveNext
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
oFCKeditor.Value	= rs("p_content")
oFCKeditor.Create "p_content"
%></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Sayfayi Güncelle"></td>
	</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
'################################################################################################################################
Sub update

id = Request.QueryString("p_id")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM PAGES WHERE p_id = "&id&""
ent.open entSQL,Connection,1,3

p_title = Request.Form("p_title")
p_content = Request.Form("p_content")
p_cat = Request.Form("p_cat")
image = Request.Form("image")
align = Request.Form("align")

ent("p_title") = p_title
ent("p_content") = p_content
ent("p_cat") = p_cat
ent("resim") = image
ent("align") = align
ent.Update
Response.Redirect "?action=hepsi"
ent.Close
Set ent = Nothing
End Sub
'################################################################################################################################
Sub del
id = Request.QueryString("p_id")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM PAGES WHERE p_id = "&id&""
delete.open delSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>
