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
If Session("Level") = "1" OR Session("Level") = "2" Then
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
ElseIF act = "butundosya" Then
call butundosyaa
End If

Sub butundosyaa
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM DOWNLOADS ORDER BY pid DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Dosya YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM DOSYALAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Dosya Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eklenme</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Indirilme Sayisi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Boyut</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Puantaj</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set cat = Connection.Execute("Select * FROM DOWN_CATS where cid="&rs("cid")&"")
%>
	<tr onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
		<td align="center"><%=rs("pid")%></td>
		<td><a href="?action=organize&pid=<%=rs("pid")%>"><%=rs("p_name")%></a></td>
		<td><A HREF="?action=all&id=<%=cat("cid")%>"><%=cat("c_name")%></A></td>
		<td align="center"><%=rs("p_date")%></td>
		<td align="center"><%=rs("p_hit")%></td>
		<td align="center"><%=rs("p_size")%></td>		
		<td align="center"><%=rs("point")%></td>				
		<td align="center"><%If rs("onay") = True then%><a href="?action=change&op=passive&pid=<%=rs("pid")%>">Pasifleþtir</a> <%else%><a href="?action=change&op=active&pid=<%=rs("pid")%>"><font color="red"><B>Aktifleþtir</B></font></a><%End if%>	- <a href="?action=organize&pid=<%=rs("pid")%>">Düzenle</a> - <a onClick="return submitConfirm18(this)" onClick="return submitConfirm18(this)" href="?action=delete&pid=<%=rs("pid")%>">Sil</a></td>
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
Connection.Execute("DELETE FROM DOWN_CATS WHERE cid = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub cat_newreg
f_1 = Request.Form("c_name")
f_2 = Request.Form("cat_info")
f_3 = Request.Form("cat_image")
If f_3=""Then
f_3="yok.gif"
End if
IF f_1="" OR f_2="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlarý Doldurunuz</div>"
ElseIF f_1<>"" AND f_2<>"" Then
IF Request.QueryString("x") = "Update" Then
Connection.Execute("UPDATE DOWN_CATS Set c_name = '"&f_1&"' WHERE cid = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE DOWN_CATS Set cat_info = '"&f_2&"' WHERE cid = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE DOWN_CATS Set cat_image = '"&f_3&"' WHERE cid = "&Request.QueryString("id")&"")
ELSE
Connection.Execute("INSERT INTO DOWN_CATS (c_name, cat_info, cat_image) VALUES ('"&f_1&"', '"&f_2&"', '"&f_3&"')")
End IF
Response.Redirect "?action=Cats"
End IF
End Sub
Sub cat_new
opt = Request.QueryString("x")
IF opt = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM DOWN_CATS WHERE cid = "&Request.QueryString("id")&"",Connection,3,3
n_page = "Update"
c_name = rs("c_name")
cat_info = rs("cat_info")
cat_image = rs("cat_image")
End IF
%>
<table border="0" cellspacing="2" cellpadding="2"  align="center" class="td_menu" width="100%">
<form method="post" action="?action=Cat_Register&x=<%=n_page%>&id=<%=Request.QueryString("id")%>">
<tr >
<td>Kategori Adý : </td>
<td ><input type="text" name="c_name" class="form1" size="75"  value="<%=c_name%>"></td>
</tr>
<tr >
<td>Kategori Açýklama : </td>
<td ><input type="text" name="cat_info" class="form1" size="75" value="<%=cat_info%>"></td>
</tr>
<tr >
<td>Kategori Resim : </td>
<td>uploads/image/Dosya_Kategorileri/<input type="text" name="cat_image" size="42" class="form1" value="<%=cat_image%>"></td>
</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Kategoriyi Kaydet"></td>
	</tr>

</form>
</table>
<%
End Sub
Sub cats
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM DOWN_CATS ORDER BY c_name ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="3" align="center"><B>-=- DOSYA KATEGORÝLERÝ -=-</B></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori Açýklamasý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<% Do While Not rs.Eof %>
	<tr onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
		<td><a href="?action=all&id=<%=rs("cid")%>"><%=rs("c_name")%></a></td>
		<td><%=rs("cat_info")%></td>
		<td align="center"><a href="?action=Cat_New&x=Edit&id=<%=rs("cid")%>">Düzenle</a> - <a onClick="return submitConfirm18(this)" href="?action=Cat_Delete&id=<%=rs("cid")%>">Sil</a></td>
		
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
SQL = "Select * FROM DOWNLOADS WHERE cid = "&Request.QueryString("id")&" AND onay = True ORDER BY pid DESC"
rs.open SQL,Connection,1,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
<%
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Bu Kategoride Yayýnlanan Eklenmiþ Dosya YOK !</div>"
ELSE
%>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Dosya Adý</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eklenme</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Indirilme Sayisi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Boyutu</td>		
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Puantaj</td>				
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%Do While Not rs.Eof%>
	<tr onMouseOver="this.style.backgroundColor='<%=forum_bg_2%>'" onMouseOut=this.style.backgroundColor="">
		<td><%=rs("pid")%></td>
		<td><a href="?action=organize&pid=<%=rs("pid")%>"><%=rs("p_name")%></a></td>
		<td align="center"><%=rs("p_date")%></td>
		<td align="center"><%=rs("p_hit")%></td>
		<td align="center"><%=rs("p_size")%></td>		
		<td align="center"><%=rs("point")%></td>				
		<td align="center"><a href="?action=change&op=passive&pid=<%=rs("pid")%>">Pasifleþtir</a> - <a href="?action=organize&pid=<%=rs("pid")%>">Düzenle</a> - <a onClick="return submitConfirm18(this)" href="?action=delete&pid=<%=rs("pid")%>">Sil</a></td>
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
Set n_cats = Connection.Execute("Select * FROM DOWN_CATS ORDER BY c_name ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.p_name.value == "") {alert("Lütfen Dosyanýn Adýný Yazýnýz... !"); return false; }
if (form.cid.value == "0") {alert("Lütfen Dosyanýn Kategorisini Seçiniz... !"); return false; }
if (form.p_size.value == "") {alert("Lütfen Dosyanýn Boyutunu Kýlobyte cýnsýnden yazýnýz... !"); return false; }
if (form.p_url.value == "<%=sett("site_adres")%>/uploads/dosyalar/") {alert("Lütfen Dosyanýn Indýrýlme adresýný tam olarak yazýnýz... !"); return false; }
if (form.p_author.value == "") {alert("Lütfen Dosyanin telif bilgilerini Yazýnýz... !"); return false; }
if (form.anahtar.value == "") {alert("Lütfen arama motorlari icin tanitici key (Anahtar) kelimeleri Yazýnýz... !"); return false; }
if (form.p_info.value == "") {alert("Lütfen Dosyayý tanýtýcý Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Dosyanýn Adý</td>
		<td>: <input type="text" name="p_name" class="form1" size="60"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="cid" size="1" class="form1">
		<option value="0" selected>Lütfen Seçiniz</option>
		<%Do Until n_cats.Eof%>
		<option value="<%=n_cats("cid")%>"><%=n_cats("c_name")%></option>
		<%
		n_cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td>Dosyanýn Boyutu</td>
		<td> : <input type="text" name="p_size" class="form1" size="8">&nbsp;&nbsp;&nbsp;Dosyanýn Adresi :  : <input type="text" name="p_url" class="form1" size="55" value="<%=sett("site_adres")%>/uploads/dosyalar/"></td>
		<td>Telif</td>
		<td> : <input type="text" name="p_author" class="form1" size="24"></td>
	</tr>
	<tr>
		<td>Anahtar Kelimeler</td>
		<td> : <input type="text" name="anahtar" class="form1" size="100%"></td>
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
		oFCKeditor.Create "p_info"
		%></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Dosyayi Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub
Sub reg

p_name = Request("p_name")
cid = Request("cid")
p_info = Request("p_info")
p_author = Request("p_author")
p_size = Request("p_size")
p_url = Request("p_url")
anahtar = Request("anahtar")
uyelik = Request("uyelik")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM DOWNLOADS",Connection,3,3
Session.LCID = 2048
ent.AddNew
ent("p_name") = p_name
ent("cid") = cid
ent("p_info") = p_info
ent("ekleyen") = Session("Member")
ent("p_date") = date()
ent("onay") = True
ent("point") = 0
ent("p_author") = p_author
ent("p_url") = p_url
ent("p_size") = p_size
ent("anahtar") = anahtar
ent("p_hit") = 0
ent("gk") = Session("gk")
ent("uyelik") = uyelik
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=all&id="&cid&""

End Sub
Sub change

operation = Request.QueryString("op")
id = Request.QueryString("pid")

If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If

Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM DOWNLOADS WHERE pid = "&id&""
chn.open chnSQL,Connection,1,3

chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")

chn.Close
Set chn = Nothing

End Sub
Sub organize

id = Request.QueryString("pid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM DOWNLOADS WHERE pid = "&id&""
rs.open SQL,Connection,1,3

Set s_cats = Connection.Execute("Select * FROM DOWN_CATS ORDER BY c_name ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.p_name.value == "") {alert("Lütfen Dosyanýn Adýný Yazýnýz... !"); return false; }
if (form.cid.value == "0") {alert("Lütfen Dosyanýn Kategorisini Seçiniz... !"); return false; }
if (form.p_size.value == "") {alert("Lütfen Dosyanýn Boyutunu Kýlobyte cýnsýnden yazýnýz... !"); return false; }
if (form.p_url.value == "<%=sett("site_adres")%>/uploads/dosyalar/") {alert("Lütfen Dosyanýn Indýrýlme adresýný tam olarak yazýnýz... !"); return false; }
if (form.p_author.value == "") {alert("Lütfen Dosyanin telif bilgilerini Yazýnýz... !"); return false; }
if (form.anahtar.value == "") {alert("Lütfen arama motorlari icin tanitici key (Anahtar) kelimeleri Yazýnýz... !"); return false; }
if (form.p_info.value == "") {alert("Lütfen Dosyayý tanýtýcý Metni Yazýnýz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&pid=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Dosyanýn Adý</td>
		<td>: <input type="text" name="p_name" class="form1" size="60"  value="<%=Oku(rs("p_name"))%>"></td>
		<td>Kategorisi</td>
		<td>: 
		<select name="cid" size="1" class="form1">
		<%
		Do Until s_cats.Eof
		IF rs("cid") = s_cats("cid") Then
		opt = "selected"
		ELSe
		opt = ""
		End IF
		Response.Write "<option value="""&s_cats("cid")&""" "&opt&">"&s_cats("c_name")&"</option>"
		s_cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td>Dosyanýn Boyutu</td>
		<td> : <input type="text" name="p_size" class="form1" size="5" value="<%=Oku(rs("p_size"))%>"> Kb.&nbsp;&nbsp;&nbsp;Dosyanýn Adresi :  : <input type="text" name="p_url" class="form1" size="55" value="<%=Oku(rs("p_url"))%>"></td>
		<td>Telif</td>
		<td> : <input type="text" name="p_author" class="form1" size="24" value="<%=Oku(rs("p_author"))%>"></td>
	</tr>
	<tr>
		<td>Anahtar Kelimeler</td>
		<td> : <input type="text" name="anahtar" class="form1" size="100%" value="<%=Oku(rs("anahtar"))%>"></td>
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
		oFCKeditor.Value	= rs("p_info")
		oFCKeditor.Create "p_info"
		%></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Dosyayi Kaydet"></td>
	</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing
End Sub
Sub update

id = Request.QueryString("pid")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM DOWNLOADS WHERE pid = "&id&""
ent.open entSQL,Connection,1,3

p_name = Request.Form("p_name")
cid = Request.Form("cid")
p_size = Request.Form("p_size")
p_url = Request.Form("p_url")
p_author= Request.Form("p_author")
anahtar = Request.Form("anahtar")
p_info = Request.Form("p_info")
uyelik = Request.Form("uyelik")
ent("p_name") = p_name
ent("cid") = cid
ent("p_size") = p_size
ent("p_url") = p_url
ent("p_author") = p_author
ent("anahtar") = anahtar
ent("p_info") = p_info
ent("uyelik") = uyelik
ent.Update
Response.Redirect "?action=all&id="&ent("cid")&""


ent.Close
Set ent = Nothing

End Sub
Sub del

id = Request.QueryString("pid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM DOWNLOADS WHERE pid = "&id&""
delete.open delSQL,Connection,1,3

Set del_comments = Server.CreateObject("ADODB.RecordSet")
delCSQL = "DELETE * FROM COMMENTS WHERE nid = "&id&""
del_comments.open delCSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>
