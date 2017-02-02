<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%call ORTA%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr>
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=top_menu5%></B></CENTER></td>
                </tr>
                <tr>
<td align="center" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<!--#include file="inc/adsense.asp" -->
<%
act = Request.QueryString("action")
If act = "members" Then
call mems
ElseIf act = "member_details" Then
call mem_details
ElseIf act = "add_friendlist" Then
call addfl
ElseIF act = "OnlineUsers" Then
call online_users
End If

Sub online_users
Session.LCID = sett("s_lcid")-22
liste = DateAdd("n", -1 * 1, Now())
Set ouRs = Server.CreateObject("ADODB.RecordSet")
ouRs.open "Select * FROM MEMBERS WHERE son_tarih >= #"&liste&"# ORDER BY son_tarih DESC",Connection,3,3
'Session.LCID = sett("s_lcid")
%>
<table border="0" cellspacing="1" cellpadding="2" width="100%" align="center" class="td_menu" bgcolor="<%=content_bg%>">
<tr height="20" class="td_menu">
<td width="20%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif"><%=OU_UserName%></td>
<td width="30%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif"><%=OU_LastSession%></td>
<td width="20%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif"><%=OU_Browser%></td>
<td width="30%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif"><%=OU_Workstation%></td>
</tr>
<%
Do Until ouRs.EoF
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = menu_color
Case Else
color = content_bg
End Select
%>
<tr bgcolor="<%=color%>">
<td width="20%" align="center"><%=ouRs("kul_adi")%></td>
<td width="30%" align="center"><%=ouRs("son_tarih")%></td>
<td width="30%" align="center"><%=ouRs("u_browser")%></td>
<td width="20%" align="center"><%=ouRs("u_ws")%></td>
</tr>
<%
ouRs.MoveNext
Loop
%>
</table>
<%
ouRs.Close
Set ouRs = Nothing
End Sub

Sub addfl
If session("enter") = "" Then
Response.Write sett("np_msg")
Else
uid = Request.QueryString("id")
uid = QS_CLEAR(uid)
If uid="" Then
Response.Redirect Request.ServerVariables("HTTP_REFERER")
Else
Set chkF = Connection.Execute("Select * FROM FRIENDS WHERE MEMBER = "&Session("Uye_ID")&" AND FRIEND = "&uid&"")
If chkF.Eof Then
Set nf = Server.CreateObject("ADODB.RecordSet")
nfSQL = "Select * FROM FRIENDS"
nf.open nfSQL,Connection,3,3

nf.AddNew
nf("MEMBER") = Session("Uye_ID")
nf("FRIEND") = uid
nf.Update
Response.Write "<BR><BR><BR><font class=td_menu><b>"&add_fl_success&"</b></font><BR><BR><BR>"
Else
Response.Write "<BR><BR><BR><font class=td_menu><b>"&added_once&"</b></font><BR><BR><BR>"
End If
%>
<!--#include file="inc/geri.asp" -->
<BR><BR><BR>
<%
End If
End If
End Sub

Sub mems

op = Request.QueryString("x")
op = QS_CLEAR(op)
IF op = "Arrange" Then
rsSQL = "Select * FROM MEMBERS WHERE Left(kul_adi,1) = '"&QS_CLEAR(Request.QueryString("y"))&"' ORDER BY kul_adi ASC"
ElseIF op = "All" OR op = "" Then
rsSQL = "Select * FROM MEMBERS ORDER BY kul_adi ASC"
End IF
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open rsSQL,Connection,3,3

Page = Request.QueryString("Page")
Page = QS_CLEAR(Page)
If Page="" Then
Page = "1"
End If
%>
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu" bgcolor="<%=tablo_cerceve%>">
<tr bgcolor="<%=content_bg%>">
<td align="center">
<%
	Response.Write "<a href='?action=members&x=All'>"&m_all&"</a>"
	For b = 65 To 90 Step 1
	Response.Write "&nbsp;-&nbsp;<a href='?action=members&x=Arrange&y=" & Chr(b) &"'>" & Chr(b) &"</a>"
	Next
%>
- <A HREF="?action=members&x=Arrange&y=Ç">Ç</A> - <A HREF="?action=members&x=Arrange&y=Ý">Ý</A> - <A HREF="?action=members&x=Arrange&y=Ö">Ö</A> - <A HREF="?action=members&x=Arrange&y=Þ">Þ</A> - <A HREF="?action=members&x=Arrange&y=Ü">Ü</A></td>
</tr>
</table>
<table width="100%" border="0" cellspacing="3" cellpadding="3" align="center" class="td_menu" style="<%=TableBorder%>">
	<tr>
		<td background=images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif>&nbsp;<B>Kullanýcý Adý</B></td>
		<td background=images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif>&nbsp;<B>Sehir</B></td>
		<td background=images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif>&nbsp;<B>Son Ziyaret</B></td>
		<td background=images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif>&nbsp;<B>Ýmzasý</B></td>
 		<td background=images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif>&nbsp;<B>Derecesi</B></td>
	</tr>
<%
IF rs.Eof OR rs.Bof Then
Response.Write ""
ELSE
rs.pagesize = "50"
rs.absolutepage = Page
sayfa = rs.PageCount
For ar = 1 To rs.pagesize
If rs.eof Then Exit For
IF UCase(Left(rs("kul_adi"),1)) <> UCase(QS_CLEAR(Request.QueryString("y"))) AND op = "Arrange" Then
Response.Write ""
ELSE
percent = rs("msg_sayisi")+rs("login_sayisi")/100
If percent >= 100 Then
percent = "100"
Else
percent = percent
End If
%>
	<tr onMouseOver="this.style.backgroundColor='<%=background%>'" onMouseOut=this.style.backgroundColor="">
		<td>&nbsp;<a href="?action=member_details&uid=<%=rs("uye_id")%>"><%=rs("kul_adi")%></a></td>
		<td>&nbsp;<%=rs("sehir")%></td>
		<td>&nbsp;<%=rs("son_tarih")%></td>
		<td>&nbsp;<%=rs("imza")%></td>
		<td><img src="images/ankett.gif" width="<%=percent%>" height="20"><IMG SRC="images/anket.gif"></td>		
	</tr>
<%
End IF
rs.MoveNext
Next
End IF
%>
</table>
<%
bp = Page-1
IF Page <> 1 Then
With Response
		IF QS_CLEAR(Request.QueryString("x")) <> "Arrange" Then
	.Write "<a href=""?action=members&Page="&bp&""">« "
		ELSE
	.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"))&"&Page="&bp&""">« "
		End IF
	.Write previous_page & "</a>"
	.Write "&nbsp;-&nbsp;"
End With
End IF
for y=1 to sayfa
IF s=y then
Response.Write y
ELSE
With Response
		IF QS_CLEAR(Request.QueryString("x")) <> "Arrange" Then
	.Write "<a href=""?action=members&Page="&y&""">"
		ELSE
	.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"))&"&Page="&y&""">"
		End IF
	.Write "<font class=""td_menu"">["&y&"]</font></a>"
End With
End IF
Next
np = Page+1
IF NOT rs.Eof Then
With Response
	.Write "&nbsp;-&nbsp;"
		IF QS_CLEAR(Request.QueryString("x")) <> "Arrange" Then
	.Write "<a href=""?action=members&Page="&np&""">"
		ELSE
	.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"))&"&Page="&np&""">"
		End IF
	.Write next_page & " »</a>"
End With
End IF

End Sub
Sub mem_details
If Session("Enter") = "" Then
Response.Write WriteMsg(no_entry)
else
uyeid = Request.QueryString("mid")
uyeid = QS_CLEAR(uyeid)
IF uyeid="" Then
uyeid = Request.QueryString("uid")
uyeid = QS_CLEAR(uyeid)
END IF

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MEMBERS WHERE uye_id = "&uyeid&""
rs.open SQL,Connection,3,3

IF rs.EoF Then
Response.Write WriteMsg(error17)
ELSE
'Session.LCID = S_LCID
kadi = rs("kul_adi")
name = rs("isim")
login = rs("login_sayisi")
lastlogin = rs("son_tarih")
regdate = rs("uyelik_tarihi")
msg_count = rs("msg_sayisi")

If rs("sehir") = "" Then
city = ""&detail_empty&""
Else
city = rs("sehir")
End If

If rs("meslek") = "" Then
job = ""&detail_empty&""
Else
job = rs("meslek")
End If

If rs("cinsiyet") = "" Then
sex = ""&detail_empty&""
ElseIf rs("cinsiyet") = "a" Then
sex = ""&male&""
ElseIf rs("cinsiyet") = "b" Then
sex = ""&female&""
Else
sex = rs("cinsiyet")
End If

strAge = Int(Right(rs("yas"),4))
strNow = Int(Right(Date(),4))

age = strNow - strAge

If rs("imza") = "" Then
signature = detail_empty
Else
signature = rs("imza")
End IF

IF rs("seviye") = "1" Then
seviye = level1
ElseIF rs("seviye") = "2" Then
seviye = level2
ElseIF rs("seviye") = "3" Then
seviye = level3
ElseIF rs("seviye") = "4" Then
seviye = level4
ElseIF rs("seviye") = "5" Then
seviye = level5
ELSE
seviye = normal
End IF
%>
<center><b><%=rs("kul_adi")%></b> (<%=seviye%>)</center>
<table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu" bgcolor="<%=tablo_cerceve%>">
<tr bgcolor="<%=content_bg%>">
<td valign="top">
<table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu">
<tr align="center" height="18" bgcolor="<%=content_bg%>">
<td colspan="2" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif"><%=detail_avatar%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td colspan="2" align="center">
<%
If rs("resmim_onay")=true Then
resim=rs("resmim")
Else
resim="yok.gif"
End if
%>
<IMG SRC="uploads/avatar/<%=resim%>" WIDTH="120" BORDER="0">
<br>
<%
IF rs("u_online") = True Then
detail_user_image = "user_online.gif"
ELSE
detail_user_image = "user_offline.gif"
End IF
%>
<img src="images/temalar/<%=Session("SiteTheme")%>/<%=detail_user_image%>" border="0">
</td>
</tr>

</table>
</td>
<td valign="top">
<table border="0" cellspacing="1" cellpadding="0" width="100%" class="td_menu">
<tr align="center" height="18" bgcolor="<%=content_bg%>">
<td colspan="2" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif"><%=detail_about%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_city%> : </td>
<td width="50%" align="left"><%=city%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_job%> : </td>
<td width="50%" align="left"><%=job%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_sex%> : </td>
<td width="50%" align="left"><%=sex%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_age%> : </td>
<td width="50%" align="left"><%=age%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_logcount%> : </td>
<td width="50%" align="left"><%=login%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_msgcount%> : </td>
<td width="50%" align="left"><%=msg_count%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_regdate%> : </td>
<td width="50%" align="left"><%=rs("uyelik_tarihi")%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right"><%=detail_lastlogin%> : </td>
<td width="50%" align="left"><%=rs("son_tarih")%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td width="50%" align="right" valign="top"><%=detail_signature%> : </td>
<td width="50%" align="left"><%=signature%></td>
</tr>
</table>
</td>
</tr>
</table>
<%
If Session("Enter") = "Yes" Then
If Session("Member") <> rs("kul_adi") Then
Set chkF = Connection.Execute("Select * FROM FRIENDS WHERE MEMBER = "&Session("Uye_ID")&" AND FRIEND = "&QS_CLEAR(Request.QueryString("uid"))&"")
IF NOT chkF.EoF Then
addfl_button = " disabled"
End IF
%>
<br>
<table border="0" cellspacing="1" cellpadding="1" align="right">
<tr>
<form method="post" action="msgbox.asp?action=new_msg&who=<%=rs("kul_adi")%>">
<td align="right">
<input type="submit" value="<%=msg_send%>" class="submit">
</td>
</form>
<form method="post" action="?action=add_friendlist&id=<%=rs("uye_id")%>">
<td align="left">
<input type="submit" value="Arkadaþ Listeme Ekle" class="submit" <%=addfl_button%>>
</td>
</form>
</tr>
</table>
<%
End IF
End IF
End IF
End IF
End Sub
%>
<!--#include file="inc/adsense.asp" -->
</td>
</tr>
</table>
<%
call ORTA2
call ALT
%>