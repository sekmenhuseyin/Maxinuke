<!--#include file="inc/includes.asp" -->
<!--#include file="guvenlik.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<link rel="stylesheet" type="text/css" href="images/temalar/<%=Session("SiteTheme")%>/style.css">
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
IF Session("Level") = "1" OR Session("Level") = "5" Then
act = Request.QueryString("eylem")
act = QS_CLEAR(act)
If act = "ayarlar" Then
call ayarlarr
ElseIf act = "Updatef_ayarlar" Then
call update_f_ayarlar
ElseIf act = "katekle" Then
call kateklee
ElseIf act = "RegCat" Then
call cat_reg
ElseIf act = "yeniforum" Then
call yeniforumm
ElseIf act = "RegForum" Then
call freg


ElseIf act = "EditCat" Then
call cat_edit
ElseIf act = "UpdateCat" Then
call upd_cat
ElseIf act = "DeleteCat" Then
call del_cat
ElseIf act = "Lock" Then
call lock
ElseIf act = "UnLock" Then
call unlock
ElseIf act = "SetEntry" Then
call entry
ElseIf act = "SetNoEntry" Then
call noentry
ElseIf act = "UpdateMsg" Then
call updmsg
ElseIf act = "DeleteMsg" Then
call delmsg
ElseIf act = "LockMsg" Then
call msglock
ElseIf act = "UnLockMsg" Then
call msgunlock
ElseIF act = "Edit-Rep" Then
call rep_edit
ElseIF act = "Update-Rep" Then
call rep_update
ElseIF act = "Del-Rep" Then
call rep_delete
End If
If Request.QueryString("noldu")="oldu" then
%>
<div align="center" class="td_menu"><B>Ýþlem baþarýyla tamamlandý.</B></div>
<%
End if
'#############################################################################################################################################
Sub ayarlarr
Set sRs = Server.CreateObject("ADODB.RecordSet")
sRs.open "Select * FROM SETTINGS",Connection,3,3
%>
<form action="?eylem=Updatef_ayarlar" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
	<tr>
		<td>Bir Sayfada Gösterilecek Konu Sayýsý : </td>
		<td width="20%" >
		<select size="1" name="topics" class="form1">
		<%
		For I = 1 TO 50
		IF sRs("f_posts") = I Then
		opt = "selected"
		Else
		opt = ""
		End IF
		%>
		<option value="<%=I%>" <%=opt%>><%=I%></option>
		<% Next %>
		</select>
		</td>
	</tr>
	<tr>
		<td>Bir Sayfada Gösterilecek Yanýt Sayýsý : </td>
		<td width="20%" >
		<select size="1" name="replies" class="form1">
		<%
		For I = 1 TO 50
		IF sRs("f_replies") = I Then
		opt = "selected"
		Else
		opt = ""
		End IF
		%>
		<option value="<%=I%>" <%=opt%>><%=I%></option>
		<% Next %>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="Tamam »" class="submit" style="width:100%"></td>
	</tr>
</table>
</form>
<%
End Sub
'##############################################################################################################################################
Sub update_f_ayarlar
Connection.Execute("UPDATE SETTINGS SET f_posts = "&Request.Form("topics")&"")
Connection.Execute("UPDATE SETTINGS SET f_replies = "&Request.Form("replies")&"")
Response.Redirect "?eylem=ayarlar&noldu=oldu"
End Sub
'##############################################################################################################################################
Sub kateklee
%>
<form action="?eylem=RegCat" method="post">
<table border="0" cellspacing="0" cellpadding="1" align="center" class="td_menu">
	<tr>	
		<td>Yeni Ana Kategori Adý : </td>
		<td><input type="text" name="katad" class="form1" size="30"></td>
	</tr>
	<tr>	
		<td>Yeni Ana Kategori Sirasi : </td>
		<td>
		<SELECT NAME="sira">
			<OPTION VALUE="1">1
			<OPTION VALUE="2">2
			<OPTION VALUE="3">3
			<OPTION VALUE="4">4
			<OPTION VALUE="5">5
			<OPTION VALUE="6">6
			<OPTION VALUE="7">7
			<OPTION VALUE="8">8
			<OPTION VALUE="9">9
			<OPTION VALUE="10">10
		</SELECT>
		</td>
	</tr>
	<tr>
		<td colspan="2"><input type="submit" value="K A Y D E T" class="submit" style="width:100%"></td>
	</tr>
</table>
</form>
<%
End Sub
'##############################################################################################################################################
Sub cat_reg
name = duz(Request.Form("katad"))
sira = duz(Request.Form("sira"))
If name="" Then
Response.Write "<center><font face=verdana size=1>Lütfen Kategori Adý Yazýnýz...</font></center>"
%><BR>
<CENTER><!--#include file="inc/geri.asp" --><BR><BR></CENTER>
<%
ELSE
Set cent = Server.CreateObject("ADODB.RecordSet")
cSQL = "Select * FROM f_anakategori"
cent.open cSQL,Connection,3,3
cent.AddNew
cent("anakatadi") = name
cent("sira") = sira
cent.Update
Response.Redirect "forum.asp"
End IF
End Sub
'##############################################################################################################################################
Sub yeniforumm
Set cats = Connection.Execute("Select * FROM f_anakategori")
%>
<form action="?eylem=RegForum" method="post">
<table border=0 cellspacing=2 cellpadding=2 align=center class="td_menu">
	<tr>
		<td>Forum Adý : </td>
		<td><input type="text" name="fname" size="50" class="form1"></td>
	</tr>
	<tr>
		<td>Açýklamasý : </td>
		<td><textarea name="finfo" cols="50" rows="5" class="form1"></textarea></td>
	</tr>
	<tr>
		<td>Kategori : </td>
		<td>
		<select name="fcat" size="1" class="form1">
		<option value="0" selected>--------- Lutfen Forumun Kategorisini Seciniz ---------</option>
		<% Do While NOT cats.Eof %>
		<option value="<%=cats("anakatid")%>"><%=cats("anakatadi")%></option>
		<%
		cats.MoveNext
		Loop
		%>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="2"><input type="submit" value="K A Y D E T" class="submit" style="width: 100%"></td>
	</tr>
</table>
</form>
<%
End Sub
'##############################################################################################################################################
Sub freg
namef = duz(Request.Form("fname"))
infof = duz(Request.Form("finfo"))
catf = duz(Request.Form("fcat"))
If namef="" OR infof="" OR catf="0" Then
Response.Write "<center><font face=verdana size=1>Lütfen Tüm Alanlarý Doldurunuz...</font></center>"
%>
<CENTER><BR><!--#include file="inc/geri.asp" --><BR><BR></CENTER>
<%
ELSE
Set creg = Server.CreateObject("ADODB.RecordSet")
cregSQL = "Select * FROM f_kategoriler"
creg.open cregSQL,Connection,3,3
creg.AddNew
creg("katad") = namef
creg("kataciklama") = infof
creg("anakatid") = catf
creg.Update
creg.Close
Set creg = Nothing
Response.Redirect "forum.asp"
End IF
End Sub




Sub cat_edit

Set rs = Connection.Execute("Select * FROM f_kategoriler WHERE kategoriid = "&QS_CLEAR(Request.QueryString("id"))&"")
Set cats = Connection.Execute("Select * FROM f_anakategori")
%>
<form action="?eylem=UpdateCat&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center"  style="font-family:verdana; font-size:11px;">
<tr>
<td  >Forum Adý : </td>
<td  ><input type="text" name="katad" class="form1" size="70" value="<%=rs("katad")%>"></td>
</tr>
<tr>
<td  >Forum Açýklamasý : </td>
<td  ><textarea name="kataciklama" rows="10" cols="70" class="form1"><%=rs("kataciklama")%></textarea></td>
</tr>
<tr>
<td  >Forum Kategori : </td>
<td  >
<select name="mcat" size="1" class="form1">
<%
Do While NOT cats.Eof
If rs("anakatid") = cats("anakatid") Then
sel = "selected"
Else
sel = ""
End If
%>
<option value="<%=cats("anakatid")%>" <%=sel%>><%=cats("anakatadi")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr>
<td  >&nbsp;</td>
<td  ><input type="submit" class="submit" value="G ü n c e l l e"></td>
</tr>
</table>
</form>
<%
rs.Close
Set rs = Nothing

End Sub
Sub upd_cat
id = Request.QueryString("id")
id = QS_CLEAR(id)
name = duz(Request.Form("katad"))
info = duz(Request.Form("kataciklama"))
cat = duz(Request.Form("mcat"))
Set cupd = Server.CreateObject("ADODB.RecordSet")
cupdSQL = "Select * FROM f_kategoriler WHERE kategoriid = "&id&""
cupd.open cupdSQL,Connection,3,3

If NOT cupd.Eof Then
cupd("katad") = name
cupd("kataciklama") = info
cupd("anakatid") = cat
cupd.Update
END IF
cupd.Close
Set cupd = Nothing
Response.Redirect "forum.asp"
End Sub
Sub unlock
fid = Request.QueryString("id")
fid = QS_CLEAR(fid)
Set upd = Server.CreateObject("ADODB.RecordSet")
updSQL = "SELECT * FROM f_kategoriler WHERE kategoriid = "&fid&""
upd.open updSQL,Connection,3,3

If NOT upd.Eof Then
upd("kilit") = False
upd.Update
End IF
upd.Close
Set upd = Nothing
Response.Redirect "forum.asp"

End Sub



Sub del_cat
id = Request.QueryString("id")
id = QS_CLEAR(id)
Set del1 = Server.CreateObject("ADODB.RecordSet")
d1SQL = "DELETE FROM f_kategoriler WHERE kategoriid = "&id&""
del1.open d1SQL,Connection,3,3
Set del2 = Server.CreateObject("ADODB.RecordSet")
d2SQL = "DELETE FROM f_mesajlar WHERE kategoriid = "&id&""
del2.open d2SQL,Connection,3,3
Response.Redirect "forum.asp"

End Sub

Sub Lock

fid = Request.QueryString("id")
fid = QS_CLEAR(fid)
Set upd = Server.CreateObject("ADODB.RecordSet")
updSQL = "SELECT * FROM f_kategoriler WHERE kategoriid = "&fid&""
upd.open updSQL,Connection,3,3

If NOT upd.Eof Then
upd("kilit") = True
upd.Update
End IF
upd.Close
Set upd = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")

End Sub



Sub entry

kategoriid = Request.QueryString("id")
kategoriid = QS_CLEAR(kategoriid)
Set cnoentry = Connection.Execute("UPDATE f_kategoriler SET noentry = False WHERE kategoriid="&kategoriid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")

End Sub

Sub noentry
kategoriid = Request.QueryString("id")
kategoriid = QS_CLEAR(kategoriid)
Set cnoentry = Connection.Execute("UPDATE f_kategoriler SET noentry = True WHERE kategoriid="&kategoriid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub delmsg
msid = Request.QueryString("id")
msid = QS_CLEAR(msid)
Set mdel = Server.CreateObject("ADODB.RecordSet")
mdSQL = "DELETE * FROM f_mesajlar WHERE mesajid = "&msid&" AND topic = True"
mdel.open mdSQL,Connection,3,3
Set mdel2 = Server.CreateObject("ADODB.RecordSet")
md2SQL = "DELETE * FROM f_mesajlar WHERE okunma = "&msid&" AND topic = False"
mdel2.open md2SQL,Connection,3,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub msglock
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
Set msgul = Connection.Execute("UPDATE f_mesajlar SET kilit = True WHERE mesajid="&msgid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub msgunlock
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
Set msgul = Connection.Execute("UPDATE f_mesajlar SET kilit = False WHERE mesajid="&msgid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

Sub rep_edit
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM f_mesajlar WHERE mesajid = "&QS_CLEAR(Request.QueryString("id"))&"",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="50%" align="center" class="td_menu">
<form action="?eylem=Update-Rep&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<tr bgcolor="<%=menu_color%>">
<td>
<textarea name="msg" rows="20" cols="100" class="form1"><%=rs("mmsg")%></textarea>
</td>
</tr>
<tr>
<td align="center">
<input type="submit" value="Güncelle »" class="submit">
</td>
</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing
End Sub

Sub rep_update
IF Request.Form("msg") = "" Then
Response.Write "<br><center>Tüm Alanlarý Doldurunuz</center><br>"
ELSE
Set mRs = Connection.Execute("Select * FROM f_mesajlar WHERE mesajid = "&QS_CLEAR(Request.QueryString("id"))&"")
frm_mesajid = mRs("okunma")
Connection.Execute("UPDATE f_mesajlar SET mmsg = '"&Request.Form("msg")&"' WHERE mesajid = "&QS_CLEAR(Request.QueryString("id"))&"")
Response.Redirect "?eylem=Topic&id="&frm_mesajid&""
End IF
End Sub

Sub rep_delete
Set mRs = Connection.Execute("Select * FROM f_mesajlar WHERE mesajid = "&QS_CLEAR(Request.QueryString("id"))&"")
frm_mesajid = mRs("okunma")
Connection.Execute("DELETE FROM f_mesajlar WHERE mesajid = "&QS_CLEAR(Request.QueryString("id"))&"")
Response.Redirect "Forum.asp?eylem=Topic&id="&frm_mesajid&""
End Sub
End if
%>