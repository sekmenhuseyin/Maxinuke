<!--#include file="guvenlik.asp" -->
<!--#include file="inc/includes.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
If Session("Level") = "1" Then
act = Request.QueryString("action")
If act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "new_enter" Then
call enternew
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
ElseIF act = "AllNews" Then
call all_news
End If

Sub all_news
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM muzikler where tip=1 ORDER BY sanatci",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu>Eklenmiþ Muzik YOK !</div>"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu">
	<tr>
		<td colspan="8" align="center"><B>-=- TUM MUZIKLER -=-</B></td>
	</tr>
	<tr>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Ýd</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Sanatci</td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Muzik Adi </td>
		<td align="center" background="images/temalar/<%=Session("SiteTheme")%>/blok_ust_orta.gif">Eylem</td>
	</tr>
<%
Do Until rs.EoF
%>
	<tr>
		<td align="center"><%=rs("muzik_id")%></td>
		<td><%=rs("sanatci")%></td>
		<td><%=rs("parca")%></td>
		<td align="center"><%If rs("onay") = false then %><a href="?action=change&op=active&muzik_id=<%=rs("muzik_id")%>"><font color="#FF0000"><B>Aktifleþtir</B></a> <%else%><a href="?action=change&op=passive&muzik_id=<%=rs("muzik_id")%>">Pasifleþtir</a> <%End if%>- <a href="?action=organize&muzik_id=<%=rs("muzik_id")%>">Düzenle</a> - <a href="?action=delete&muzik_id=<%=rs("muzik_id")%>">Sil</a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
End Sub

Sub enternew
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.sanatci.value == "") {alert("Lütfen Sanatciyi yaziniz... !"); return false; }
if (form.parca.value == "") {alert("Lütfen Muzik Adi Yazýnýz... !"); return false; }
if (form.dosya_adi.value == "") {alert("Lütfen muzik dosyasini bilgisayarinizdan yukleyiniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=register" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Sanatci</td>
		<td>: <input type="text" name="sanatci" class="form1" size="60"></td>
	</tr>
	<tr>
		<td>Parca Adi</td>
		<td>: <input type="text" name="parca" class="form1" size="60"></td>
	</tr>
	<tr>
		<td colspan="4"><iframe src="?action=IS" width="100%" height="30" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></td>
	</tr>
	<tr>
		<td>Muzik Doyasi</td>
		<td>: <input type="text" name="dosya_adi" class="form1" size="60"></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Muziki Kaydet"></td>
	</tr>
</form>
</table>
<%
End Sub
Sub reg
sanatci = Request("sanatci")
parca = Request("parca")
dosya_adi = Request("dosya_adi")
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM muzikler",Connection,3,3
ent.AddNew
ent("sanatci") = sanatci
ent("parca") = parca
ent("dosya_adi") = dosya_adi
ent("onay") = True
ent("tip") = "1"
ent.Update
ent.Close
Set ent = Nothing
session("dosya_adi")=""
Response.Redirect "?action=AllNews"
End Sub
Sub ImageSelect
%>
<table border="0" cellspacing="0" cellpadding="2" class="td_menu">
<tr><td colspan="2"></td></tr>
<form ENCTYPE="multipart/form-data" ACTION="?action=IU&ne=<%=Request.QueryString("ne")%>" METHOD="POST">
<tr >
<td align="right">Muzik dosyasi</td>
<td> : <input NAME="File2" class="form1" SIZE="20" TYPE="file"> <input type="submit" value="Gönder »" class="submit"></td>
</tr>
<tr><td colspan="2"></td></tr>
</form>
</table>
<%
End Sub
Sub ImageUpload
ImageDir = "uploads/mp3"
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
	.Write "<font class=""td_menu"">Lutfen koyu yaziyi <B>"&session("dosya_adi")&"</B> alttaki text kutusuna yaziniz"
	.Write "</font>"
End With
	End IF
End Sub
Sub change

operation = Request.QueryString("op")
id = Request.QueryString("muzik_id")

If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If

Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM muzikler WHERE muzik_id = "&id&""
chn.open chnSQL,Connection,1,3

chn("onay") = s
chn.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")

chn.Close
Set chn = Nothing

End Sub
Sub organize

id = Request.QueryString("muzik_id")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM muzikler WHERE muzik_id = "&id&""
rs.open SQL,Connection,1,3
%>
<SCRIPT LANGUAGE=JAVASCRIPT>
function formkontrol(form) 
{
if (form.sanatci.value == "") {alert("Lütfen Muzik Metni Yazýnýz... !"); return false; }
if (form.parca.value == "") {alert("Lütfen Muzik Adi Yazýnýz... !"); return false; }
if (form.dosya_adi.value == "") {alert("Lütfen dosyasini yukleyiniz... !"); return false; }
return true;
}
</SCRIPT>
<form onSubmit="return formkontrol(this)" action="?action=update&muzik_id=<%=id%>" method="post">
<table border="0" cellspacing="5" cellpadding="5" align="center" class="td_menu" width=99%>
	<tr>
		<td>Sanatci</td>
		<td>: <input type="text" name="sanatci" class="form1" size="60" value="<%=rs("sanatci")%>"></td>
	</tr>
	<tr>		
		<td>Parca adi</td>
		<td>: <input type="text" name="parca" class="form1" size="60" value="<%=rs("parca")%>"></td>
	</tr>
	<tr>		
		<td>Muzik Dosyasi</td>
		<td>: <input type="text" name="dosya_adi" class="form1" size="60" value="<%=rs("dosya_adi")%>"></td>
	</tr>
	<tr>
		<td align="center" colspan="4"><input type="submit" class="submit"  name="submit" style="width:100%" value="Muziki Güncelle"></td>
	</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

id = Request.QueryString("muzik_id")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM muzikler WHERE muzik_id = "&id&""
ent.open entSQL,Connection,1,3
sanatci = Request.Form("sanatci")
parca = Request.Form("parca")
dosya_adi = Request.Form("dosya_adi")
ent("sanatci") = sanatci
ent("parca") = parca
ent("dosya_adi") = dosya_adi
ent.Update
Response.Redirect "?action=AllNews"
ent.Close
Set ent = Nothing
End Sub

Sub del
id = Request.QueryString("muzik_id")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM muzikler WHERE muzik_id = "&id&""
delete.open delSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>