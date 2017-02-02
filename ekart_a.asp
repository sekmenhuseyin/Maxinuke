<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Or Session("Level") = "2" Then
act = Request.QueryString("action")
If act = "all" Then
call hepsii
ElseIf act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "yeni" Then
call yeniii
ElseIf act = "yeni_ap" Then
call yenii_ap
ElseIf act = "yeni_midi" Then
call yeni_midiiii
ElseIf act = "register" Then
call reg
ElseIf act = "IS" Then
call ImageSelect
ElseIf act = "IU" Then
call ImageUpload
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
ElseIF act = "hepsi_ap" Then
call hepsii_ap
ElseIF act = "hepsi_midi" Then
call hepsii_midi
End If

Sub hepsii
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ekartlar where tip=1 ORDER BY id DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ E-Kart YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM E-KARTLAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Kategori</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Hiti</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set katid = Connection.Execute("Select * FROM ekart_kat WHERE katid = "&rs("katid")&"")
%>
	<tr>
		<td align="center"><%=rs("id")%></td>

		<td><%=katid("isim")%></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%=rs("hit")%></td>
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

Sub hepsii_ap
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ekartlar where tip=2 ORDER BY id DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Arka Plan YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM ARKA PLANLAR -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Resim</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Hiti</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set katid = Connection.Execute("Select * FROM ekart_kat WHERE katid = "&rs("katid")&"")
%>
	<tr>
		<td align="center"><IMG SRC="uploads/ekart/<%=rs("buyuk")%>" WIDTH="120" HEIGHT="118" BORDER="0" ALT=""></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%=rs("hit")%></td>
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&id=<%=rs("id")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&id=<%=rs("id")%>">Pasifleþtir</a> <%End if%>- <a href="?action=delete&id=<%=rs("id")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
End Sub

Sub hepsii_midi
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ekartlar where tip=3 ORDER BY yol",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Ses dosyasi YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM EKART MUZIK DOSYALARI -=-</B></td>
	</tr>
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Adi</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Tarih</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Hiti</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
Set katid = Connection.Execute("Select * FROM ekart_kat WHERE katid = "&rs("katid")&"")
%>
	<tr>
		<td><%=rs("yol")%></td>
		<td align="center"><%=rs("tarih")%></td>
		<td align="center"><%=rs("hit")%></td>
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&id=<%=rs("id")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&id=<%=rs("id")%>">Pasifleþtir</a> <%End if%> - <a href="?action=delete&id=<%=rs("id")%>">Sil</a></td>
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
Connection.Execute("DELETE FROM ekart_kat WHERE katid = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub cat_newreg
f_1 = Request.Form("name")
IF f_1="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlarý Doldurunuz</div>"
ElseIF f_1<>"" Then
IF Request.QueryString("x") = "Update" Then
Connection.Execute("UPDATE ekart_kat Set isim = '"&f_1&"' WHERE katid = "&Request.QueryString("id")&"")
ELSE
Connection.Execute("INSERT INTO ekart_kat (isim) VALUES ('"&f_1&"')")
End IF
Response.Redirect "?action=Cats"
End IF
End Sub
Sub cat_new
opt = Request.QueryString("x")
IF opt = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ekart_kat WHERE katid = "&Request.QueryString("id")&"",Connection,3,3
n_page = "Update"
c_name = rs("isim")
btn_submit = "Güncelle"
ELSE
btn_submit = "Ekle +"
End IF
%>
<table border="0" cellspacing="2" cellpadding="2"  align="center" class="td_menu">
<form method="post" action="?action=Cat_Register&x=<%=n_page%>&id=<%=Request.QueryString("id")%>">
<tr>
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
rs.open "Select * FROM ekart_kat ORDER BY isim ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="5" align="center"><B>-=- E-Kart KATEGORÝLERÝ -=-</B></td>
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



Sub yeniii
Set n_cats = Connection.Execute("Select * FROM ekart_kat ORDER BY isim ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.yol.value == "") {alert("Lütfen Kucuk Onizleme resmini yukleyip yaziniz... !"); return false; }
if (form.buyuk.value == "") {alert("Lütfen Buyuk Ekart resmini yukleyip yaziniz ... !"); return false; }
if (form.katid.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<input type="hidden" name="tip" value="1">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Kucuk Resim URL</td>
		<td>: <input type="text" name="yol" class="form1" size="60"> (Max 95X75 Geniþlik)</td>
		<td><iframe src="?action=IS" width="100%" height="30" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></td>
	</tr>
	<tr>		
		<td>Buyuk Resim URL</td>
		<td>: <input type="text" name="buyuk" class="form1" size="60"></td>
		<td><iframe src="?action=IS" width="100%" height="30" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></td>
	</tr>
	<tr>
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
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="E-Karti Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub

Sub yenii_ap
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.buyuk.value == "") {alert("Lütfen Arkaplan resmini yukleyip yaziniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<input type="hidden" name="tip"value="2">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Arka Plan Resim URL</td>
		<td>: <input type="text" name="buyuk" class="form1" size="60"> (Max 95X75 Geniþlik)</td>
		<td><iframe src="?action=IS" width="100%" height="30" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Arka Plan Resmimi Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub

Sub yeni_midiiii
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.buyuk.value == "") {alert("Lütfen Muzik dosyasinin adini yaziniz... !"); return false; }
if (form.yol.value == "") {alert("Lütfen Midi ses dosyasini yukleyip yaziniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<input type="hidden" name="tip" value="3">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Muzik Adi</td>
		<td>: <input type="text" name="buyuk" class="form1" size="60"></td>
	</tr>
	<tr>
		<td>Ses Dosyasi URL</td>
		<td>: <input type="text" name="yol" class="form1" size="60"></td>
	</tr>
	<tr>
		<td><iframe src="?action=IS" width="100%" height="30" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Ses Dosyami Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub

Sub reg
tip = Request("tip")
katid = Request("katid")
If katid = "" then
katid = 0
End if
yol = Request("yol")
buyuk = Request("buyuk")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM ekartlar",Connection,3,3
Session.LCID = 2048
ent.AddNew
ent("katid") = katid
ent("yol") = yol
ent("buyuk") = buyuk
ent("tarih") = Now()
ent("onay") = True
ent("hit") = 0
ent("tip") = tip
ent.Update
ent.Close
Set ent = Nothing
If tip = "1" then
Response.Redirect "?action=hepsi"
elseIf tip = "2" then
Response.Redirect "?action=hepsi_ap"
elseIf tip = "3" then
Response.Redirect "?action=hepsi_midi"
End if
End Sub
Sub ImageSelect
%>
<table border="0" cellspacing="0" cellpadding="2" class="td_menu">
<tr><td colspan="2"></td></tr>
<form ENCTYPE="multipart/form-data" ACTION="?action=IU&ne=<%=Request.QueryString("ne")%>" METHOD="POST">
<tr >
<td align="right">Resim : </td>
<td><input NAME="File2" class="form1" SIZE="20" TYPE="file"> <input type="submit" value="Gönder »" class="submit"></td>
</tr>
<tr><td colspan="2"></td></tr>
</form>
</table>
<%
End Sub
Sub ImageUpload
ImageDir = "uploads/ekart"
	ForWriting = 2
	adLongVarChar = 201
	lngNumberUploaded = 0

	'Get binary data from form		
	noBytes = Request.TotalBytes 
	binData = Request.BinaryRead (noBytes)

	'convery the binary data to a string
	Set RST = CreateObject("ADODB.Recordset")
	LenBinary = LenB(binData)
	
	if LenBinary > 0 Then
	RST.Fields.Append "myBinary", adLongVarChar, LenBinary
	RST.Open
	RST.AddNew
	RST("myBinary").AppendChunk BinData
	RST.Update
	strDataWhole = RST("myBinary")
	End if

	strBoundry = Request.ServerVariables ("HTTP_CONTENT_TYPE")
	lngBoundryPos = instr(1, strBoundry, "boundary=") + 8 
	strBoundry = "--" & right(strBoundry, len(strBoundry) - lngBoundryPos)
	lngCurrentBegin = instr(1, strDataWhole, strBoundry)
	lngCurrentEnd = instr(lngCurrentBegin + 1, strDataWhole, strBoundry) - 1
	Do While lngCurrentEnd > 0
	'Get the data between current boundry and remove it from the whole.
	strData = mid(strDataWhole, lngCurrentBegin, lngCurrentEnd - lngCurrentBegin)
	strDataWhole = replace(strDataWhole, strData,"")
	
	'Get the full path of the current file.
	lngBeginFileName = instr(1, strdata, "filename=") + 10
	lngEndFileName = instr(lngBeginFileName, strData, chr(34)) 
	'Make sure they selected a file.	
	if lngBeginFileName = lngEndFileName and lngNumberUploaded = 0 Then
	Response.Redirect "index.asp?"
	End if
	'There could be an empty file box.	
	if lngBeginFileName <> lngEndFileName Then
	strFilename = mid(strData, lngBeginFileName, lngEndFileName - lngBeginFileName)

	tmpLng = instr(1, strFilename, "\")
	Do While tmpLng > 0
	PrevPos = tmpLng
	tmpLng = instr(PrevPos + 1, strFilename,"\")
	Loop
	
	FileName = right(strFilename, len(strFileName) - PrevPos)
	
	lngCT = instr(1,strData, "Content-Type:")
	
	if lngCT > 0 Then
	lngBeginPos = instr(lngCT, strData, chr(13) & chr(10)) + 4
	Else
	lngBeginPos = lngEndFileName
	End if
	lngEndPos = len(strData) 
	
	'Calculate the file size.	
	lngDataLenth = lngEndPos - lngBeginPos
	'Get the file data	
	strFileData = mid(strData, lngBeginPos, lngDataLenth)

	IF instr(1,FileName,".asp",1) Then
	Response.Write "<div align=center class=td_menu>Hatalý Dosya Türü !</div>"
	ELSE

	'Create the file.	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(server.mappath(imagedir) & "/" & FileName, ForWriting, True)
	f.Write strFileData
	Set f = nothing
	Set fso = nothing

	End IF
	
	lngNumberUploaded = lngNumberUploaded + 1
			
	End if
	
	lngCurrentBegin = instr(1, strDataWhole, strBoundry)
	lngCurrentEnd = instr(lngCurrentBegin + 1, strDataWhole, strBoundry) - 1
	loop

	IF instr(1,FileName,".asp",1) Then
	ELSE
Set wSet = Connection.Execute("Select * FROM SETTINGS")
Session("ImagePath") = FileName
wSet.Close
Set wSet = Nothing
With Response
	.Write "<font class=""td_menu"">Tebrikler Resim Baþarýyla Gönderildi. Yandaki Resim URL kutusuna baþýnda ve sonunda boþluk olmamasýna dikkat ederek <font face=arial size=4 color=red>" & Session("ImagePath") & "</font> yazýnýz."
	.Write "</font>"
End With
	End IF
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
chnSQL = "Select * FROM ekartlar WHERE id = "&id&""
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
SQL = "Select * FROM ekartlar WHERE id = "&id&""
rs.open SQL,Connection,1,3

Set s_cats = Connection.Execute("Select * FROM ekart_kat ORDER BY isim ASC")
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.yol.value == "") {alert("Lütfen Kucuk Onizleme resmini yukleyip yaziniz... !"); return false; }
if (form.buyuk.value == "") {alert("Lütfen Buyuk Ekart resmini yukleyip yaziniz ... !"); return false; }
if (form.katid.value == "0") {alert("Lütfen Kategori Seçiniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&id=<%=id%>" method="post">
<input type="hidden" name="tip" value="<%=rs("tip")%>">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Kucuk Resim URL</td>
		<td>: <input type="text" name="yol" class="form1" size="60" value="<%=rs("yol")%>"> (Max 95X75 Geniþlik)</td>
		<td><iframe src="?action=IS" width="100%" height="30" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></td>
	</tr>
	<tr>		
		<td>Buyuk Resim URL</td>
		<td>: <input type="text" name="buyuk" class="form1" size="60" value="<%=rs("buyuk")%>"></td>
		<td><iframe src="?action=IS" width="100%" height="30" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></td>
	</tr>
	<tr>
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
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="E-Karti Güncelle"></td>
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
entSQL = "Select * FROM ekartlar WHERE id = "&id&""
ent.open entSQL,Connection,1,3

katid = Request.Form("katid")
yol = Request.Form("yol")
buyuk = Request.Form("buyuk")


ent("katid") = katid
ent("yol") = yol
ent("buyuk") = buyuk
ent.Update
Response.Redirect "?action=all&id="&ent("katid")&""


ent.Close
Set ent = Nothing

End Sub
Sub del

id = Request.QueryString("id")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM ekartlar WHERE id = "&id&""
delete.open delSQL,Connection,1,3

Set del_comments = Server.CreateObject("ADODB.RecordSet")
delCSQL = "DELETE * FROM COMMENTS WHERE nid = "&id&""
del_comments.open delCSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>