<%Response.Buffer = True%>
<!--#include file="inc/includes.asp" -->
<!--#include file="bh.asp" -->
<html>
<head>
<title><%=sett("site_isim")%> - <%=sett("site_slogan")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<LINK href="images/site/favicon.ico" rel="SHORTCUT ICON">
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<META NAME="TITLE" CONTENT="<%=sett("site_isim")%> - <%=sett("site_slogan")%>">
<meta http-equiv="Content-Language" content="tr">
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="Maxinuke">
<META NAME="Description" CONTENT="<%=sett("Keywords")%>">
<META NAME="Keywords" CONTENT="<%=sett("Keywords")%>,Maxinuke">
<META NAME="Subject" CONTENT="<%=sett("Keywords")%>">
<meta name="abstract" content="<%=sett("site_isim")%>">
<META NAME="Copyright" CONTENT="Maxinuke © 2008 - 2009">
<META NAME="Distribution" CONTENT="Global">
<meta name="robots" content="index, follow, imageindex, imageclick, archive, all" >
<meta NAME="rating" content="All">
<meta name="publisher" content="<%=sett("site_adres")%>">
<meta name="audience" content="all">
<meta name="REVISIT-AFTER" content="1 days">
<meta http-equiv="reply-to" content="<%=sett("site_mail")%>">
<meta name="resource-type" content="document">
<META NAME="googlebot" CONTENT="Index, Follow">
<meta NAME="Designer" CONTENT="Maxinuke Team and Penoaydin"/>
<META HTTP-EQUIV="imagetoolbar" CONTENT="no">
<body bgcolor="<%=background%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
</head>
<div class=td_menu align=center>
<%
If Session("Enter")="" then
response.write "Bu alani kullanabilmek icin sisteme uye girisi yapmis olmaniz gerekmektedir<BR><BR><input type=button value=KAPAT onclick=javascript:parent.window.focus();top.window.close() class=submit>"
response.End
End if
If Request.QueryString("eylem") = "" Then
%>
<form action="?eylem=yukle" method="post" ENCtype="multipart/form-data"><BR>
<input type="file" name="resimadi" size="35"><BR><BR>
<input class="submit" type="submit" value=" Yükle " onClick="this.value='Yükleniyor..'">
</form>
<%
End If
If Request.QueryString("eylem") = "yukle" Then
yol = "uploads/avatar"
Randomize
Numara = INT (RND*9999999999)+1
Server.ScriptTimeout=1000
Set Upload = Server.CreateObject("Persits.Upload.1")
Upload.OverwriteFiles = False
	If Upload.SaveToMemory = 0 Then
	Response.Write "<meta http-equiv='refresh' content='0;URL=?eylem='><script>alert('Yüklenecek Bir dosya secmelisiniz !!! !!!');</script>"
	Response.End
	End If
For Each File in Upload.Files
	If File.ImageType = "UNKNOWN" Then
	Response.Write "<meta http-equiv='refresh' content='0;URL=?eylem='><script>alert('Sadece JPG, GIf ve PNG Formatli Dosyalar Yükleyebilirsiniz !!! !!!');</script>"
	File.Delete
	Response.End
	End if
Next
Set Resim = Upload.Files ("resimadi")
Adi = session("Uye_ID")&Numara&Right(Resim.FileName,4)
Path = Server.MapPath(yol & "/" & Adi)
Resim.SaveAs Path
Set Jpeg = Server.CreateObject("Persits.Jpeg" )
Jpeg.Open Path
Jpeg.Width = 120'genislik
jpeg.Height = Jpeg.OriginalHeight * Jpeg.Width / Jpeg.OriginalWidth
Jpeg.Save Path
Set resok=Server.CreateObject("Adodb.Recordset")
resok.ActiveConnection = Connection
Sorgu="SELECT * FROM MEMBERS WHERE uye_id="&session("Uye_ID")&""
resok.open Sorgu,Connection,1,3
resok.update "resmim_onay",true
resok.update "resmim",Adi
resok.close
set resok=nothing
%>
<br>
Dosya Başarıyla Kaydedildi.<BR><BR>Sayfayi yenileyerek yeni resminizi gorebilirsiniz.
<!--#include file="inc/kapat.asp" -->
<%End If%>
</div>