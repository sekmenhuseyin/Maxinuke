<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%call ORTA%>
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Mesaj islemleri Sayfa Uygulamasi V 2.1										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr>
	<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=uyemenu_msgbox%></B></CENTER></td>
</tr>
<tr>
	<td align="center" valign="top" bgcolor="<%=content_bg%>">
<%
IF Session("Enter") = "Yes" Then
Set toplamesaj = Connection.Execute("Select Count(*) as mcount FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND TO = '"&uyebilgi("kul_adi")&"'")
Set gelenmesaj = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND TO = '"&uyebilgi("kul_adi")&"'")
Set gidenmesaj = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE FROM = '"&uyebilgi("kul_adi")&"' AND MSAVE = True")
Set baba_mesaj = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 1 AND SEE = True")
	If gelenmesaj("count") >= sett("msg_limit") Then
	t_msgs = sett("msg_limit")
	Else
	t_msgs = gelenmesaj("count")
	End If
	If gidenmesaj("count") >= sett("msg_limit") Then
	gidenmesaji = sett("msg_limit")
	Else
	gidenmesaji = gidenmesaj("count")
	End If
percent = Int((toplamesaj("mcount") / sett("msg_limit")) * 100)
	If percent >= 100 Then
	percent = "100"
	Else
	percent = percent
	End If
%>
		<table border="0" cellspacing="3" cellpadding="3" width="100%"  CLASS="td_menu">
		<TR>
			<TD width="120">Kullandýðýnýz Alan : <b><%=percent%>%</b></TD>
			<TD align="center" >
				<table  width="100%" border="0" cellspacing="1" cellpadding="1"  bgcolor="<%=tablo_cerceve%>">
				<tr bgcolor="<%=content_bg%>">
					<td><img src="images/ankett.gif" width="<%=percent%>" height="14"></td>
				</tr>
				</table>
			</TD>
			<TD align="right" width="80">Boþ Alan : <b><%=100-percent%>%</b></TD>
		</TR>
		</TABLE>
		<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center" height="20" class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">
		<tr>
			<td align="center"><a href="?action=msgbox"><B><%=msg_box%> (<%=t_msgs%>)</B></a></td>
			<td align="center"><a href="?action=saved"><B><%=msg_saved%> (<%=gidenmesaji%>)</B></a></td>
			<td align="center"><a href="?action=fromadmin"><B><%=msg_admin%> (<%=baba_mesaj("count")%>)</B></a></td>
			<td align="center"><a href="?action=new_msg"><B><%=msg_new%></B></a></td>
		</tr>
		</table>
<%
action = Request.QueryString("action")
action = QS_CLEAR(action)
if action = "main" then
call messages
elseif action = "msgbox" then
call messages
elseif action = "fromadmin" then
call fromadmin
ElseIF action = "saved" Then
call saved
elseif action = "new_msg" then
call new_msg
elseif action = "read_msg" then
call read
elseif action = "msg_del" then
call delete_msg
elseif action = "fact" then
call fact
elseif action = "msg_rec" then
call msg_record
end if
'###################################################################################################################################
Sub messages
Session("Folder") = "Inbox"
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MESSAGES WHERE TO = '"&uyebilgi("kul_adi")&"' AND SEE = True AND TYPE = 0 ORDER BY mdate DESC"
rs.open SQL,Connection,3,3
%>
		<table border="0" cellspacing="0" cellpadding="0" width="100%" class="td_menu">
<%
For i = 1 TO sett("msg_limit")
If rs.Eof Then Exit For
Set fmem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs("from")&"'")
	IF fmem.Eof OR fmem.Bof Then
	msg_mid = ""
	ELSE
	msg_mid = fmem("uye_id")
	End IF
	If rs("read") = false Then
	asd="background=images/onemli.gif"
	else
	asd=""
	End if
	If fmem("resmim_onay")=true Then
	resim=fmem("resmim")
	Else
	resim="yok.gif"
	End if
%>
		<tr>
			<td <%=asd%>><a href="members.asp?action=member_details&uid=<%=msg_mid%>"><IMG SRC="uploads/avatar/<%=resim%>" WIDTH="50" BORDER="0" ALT="<%=rs("from")%>" align="left"><%=rs("from")%></a><BR><%=rs("mdate")%></td>
			<td <%=asd%>><a href="?action=read_msg&mid=<%=rs("mid")%>"><%=duz(rs("subject"))%><BR><%=Left(duz(rs("msg")),80)%>...</a></td>
			<td <%=asd%> align="center"><a href="?action=msg_del&id=<%=rs("mid")%>"><img src="images/temalar/<%=Session("SiteTheme")%>/msg_del.gif" border="0" alt="<%=msg_delete_want%>"></a></td>
		</tr>
		<tr>
			<td colspan="3"><hr></td>
		</tr>
<%
rs.MoveNext
Next
%>

		</table>
<%
Set fmem = Nothing
rs.Close
Set rs = Nothing
End Sub
'###################################################################################################################################
Sub fromadmin
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MESSAGES WHERE SEE = True AND TYPE = 1 AND MSAVE = False ORDER BY mdate DESC"
rs.open SQL,Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" align="center" bgcolor="<%=content_bg%>" class="td_menu">
	<tr bgcolor=<%=background%>>
		<td width="5%" align="center">&nbsp;</td>
		<td width="40%" align=center><B><%=msg_subject%></B></td>
		<td width="40%" align=center><B><%=msg_date%></B></td>
	</tr>
<%
Do Until rs.eof
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = forum_bg_1
Case Else
color = forum_bg_2
End Select

%>
		<tr bgcolor="<%=color%>">
			<td width="5%" align="center"><% If rs("read") = True Then %><img src="images/temalar/<%=Session("SiteTheme")%>/readed.gif"><% ElseIf rs("read") = False Then %><img src="images/temalar/<%=Session("SiteTheme")%>/not_readed.gif"><% End If %></td>
			<td width="40%"><a href="?action=read_msg&mid=<%=rs("mid")%>"><%=rs("subject")%></a></td>
			<td width="40%" align=center><%=rs("mdate")%></td>
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
'###################################################################################################################################
Sub saved
Session("Folder") = "Saved"
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MESSAGES WHERE FROM = '"&uyebilgi("kul_adi")&"' AND TYPE = 0 AND MSAVE = True ORDER BY mdate DESC"
rs.open SQL,Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" align="center" bgcolor="<%=content_bg%>" class="td_menu">
	<tr>
		<td width="5%" bgcolor=<%=background%> align="center">&nbsp;</td>
		<td width="20%" bgcolor=<%=background%> align=center><B>Gönderilen</B></td>
		<td width="40%" bgcolor=<%=background%> align=center><B><%=msg_subject%></B></td>
		<td width="40%" bgcolor=<%=background%> align=center><B><%=msg_date%></B></td>
		<td width="40%" bgcolor=<%=background%> align=center>&nbsp;</td>
	</tr>
<%
Do Until rs.EoF
Set fmem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs("from")&"'")
IF fmem.Eof OR fmem.Bof Then
msg_mid = ""
ELSE
msg_mid = fmem("uye_id")
End IF
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = forum_bg_1
Case Else
color = forum_bg_2
End Select
%>
	<tr bgcolor="<%=color%>">
		<td width="5%" align="center"><% If rs("read") = True Then %><img src="images/temalar/<%=Session("SiteTheme")%>/readed.gif"><% ElseIf rs("read") = False Then %><img src="images/temalar/<%=Session("SiteTheme")%>/not_readed.gif"><% End If %></td>
		<td width="20%" align=center><a href="members.asp?action=member_details&uid=<%=msg_mid%>"><%=duz(rs("to"))%></a></td>
		<td width="40%"><a href="?action=read_msg&mid=<%=rs("mid")%>"><%=duz(rs("subject"))%></a></td>
		<td width="40%" align=center><%=rs("mdate")%></td>
		<td width="2%" align="center"><a href="?action=msg_del&msg=saved&id=<%=rs("mid")%>"><img src="images/temalar/<%=Session("SiteTheme")%>/msg_del.gif" border="0" alt="<%=msg_delete_want%>"></a></td>
	</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
'fmem.Close
Set fmem = Nothing
rs.Close
Set rs = Nothing
End Sub
'###################################################################################################################################
Sub new_msg
msg = request.querystring("msgid")
msg = QS_CLEAR(msg)
towho = request.querystring("who")
towho = QS_CLEAR(towho)
subje = request.form("subject")
if msg<>"" then
Set mem = Connection.Execute("SELECT * FROM MESSAGES WHERE mid = "&msg&"")
msgto = mem("from")
elseif towho<>"" Then
msgto = towho
elseif towho="" OR msg="" Then
msgto = ""
end if

If subje<>"" Then
subj = "Re:"&subje&""
Else
subj = ""
End If

IF towho<>"" OR msgto<>"" Then
msg_opt1 = "checked"
ELSE
msg_opt2 = "checked"
End IF

set members = Connection.Execute("select * from MEMBERS ORDER BY kul_adi")
%>
<form action="?action=msg_rec" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu">
<tr><td width="20%" align="right" valign="top"><%=msg_to%> : </td><td width="80%"><input type="radio" name="type" value="manual" <%=msg_opt1%>>&nbsp;<input type="text" name="mem_name" class="form1" value="<%=msgto%>"><br><input type="radio" name="type" value="auto" <%=msg_opt2%>>
<select size="1" name="mem_names" class=form1>
<% Do While Not members.EOF
If members("seviye") = "1" Then
seviye = " ["&level1&"]"
ElseIf members("seviye") = "2" Then
seviye = " ["&level2&"]"
ElseIf members("seviye") = "3" Then
seviye = " ["&level3&"]"
ElseIf members("seviye") = "4" Then
seviye = " ["&level4&"]"
ElseIf members("seviye") = "5" Then
seviye = " ["&level5&"]"
ELSE
seviye = ""
End If
%>
<option value="<%=members("kul_adi")%>"><%=members("kul_adi")%><%=seviye%></option>
<% members.MoveNext
Loop %>
</select>
</td></tr>
<tr>
<td width="30%" align="right"><%=msg_subject%> : </td>
<td width="70%"><input type="text" name="subject" class="form1" value="<%=subj%>" size="40"></td>
</tr>
<tr>
<td width="30%" align="right" valign="top"><%=msg_message%> : </td>
<td width="70%"><textarea name="message" rows="8" cols="40" class="form1"></textarea></td>
</tr>
<tr>
<td width="30%" align="right" valign="top"><%=msg_save%> : </td>
<td width="70%"><input type="checkbox" name="m_save" class="form1" checked></td>
</tr>
<tr>
<td width="30%" align="right">&nbsp;</td>
<td width="70%"><input type="submit" value="<%=msg_ok%>" class="submit"></td>
</tr>
</table>
</form>
<%
members.Close
Set members = Nothing
'mem.Close
Set mem = Nothing
End Sub
'###################################################################################################################################
Sub read
id = request.querystring("mid")
id = QS_CLEAR(id)

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MESSAGES WHERE mid = "&id&""
rs.open SQL,Connection,3,3

IF uyebilgi("kul_adi") = rs("TO") Then
protect_msg = True
elseIF uyebilgi("kul_adi") = rs("from") Then
protect_msg = True
ElseIF rs("TYPE") = "1" Then
protect_msg = True
ELSE
protect_msg = False
End IF

IF protect_msg = True Then
IF Session("Folder") = "Inbox" Then
rs("read") = True
rs.Update
End IF
Set frmem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs("from")&"'")

If frmem.Eof Then
msndr = "Site Manager"
Else
msndr = "<a href=""members.asp?action=member_details&uid="&frmem("uye_id")&""">"&rs("from")&"</a>"
End If
%>
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr><td><b><%=msg_sender%> : </b><%=msndr%> // <%=rs("mdate")%></td></tr>
<tr><td><b><%=msg_subject%> : </b><%=rs("subject")%></td></tr>
<tr><td><%=SmiLey(rs("msg"))%></td></tr>
</table>
<% If rs("type") <> "1" Then %>
<table border="0" cellspacing="0" cellpadding="1" width="100" align="right">
<form method="post" action="?action=new_msg&msgid=<%=rs("mid")%>">
<tr><td><input type="hidden" name="subject" value="<%=rs("subject")%>">
<input type="submit" value="<%=reply_button%>" class="form1"></td>
</form>
<form method="post" action="?action=msg_del&id=<%=rs("mid")%>">
<td>
<input type="submit" value="<%=msg_delete_want%>" class="form1"></td></tr>
</form>
</table>
<%
End IF
End If
Set frmem = Nothing
rs.Close
Set rs = Nothing
End Sub
'###################################################################################################################################
Sub delete_msg
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
if msgid<>"" Then
Set chk_msg = Connection.Execute("Select * FROM MESSAGES WHERE mid = "&msgid&"")
	IF Session("Folder") = "Saved" Then
	Set del_msg = Connection.Execute("UPDATE MESSAGES SET msave = False WHERE mid = "&msgid&"")
	ELSE
	Set del_msg = Connection.Execute("UPDATE MESSAGES SET see = False WHERE mid = "&msgid&"")
	End IF
chk_msg.Close
Set chk_msg = Nothing
end if
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	del_msg.Close
	Set del_msg = Nothing
End Sub
'###################################################################################################################################
Sub fact
f_action = Request.Form("fr_act")
f_member = Request.Form("fr_nm")

IF f_member = "" Then
Response.Write ps_member

ElseIF f_action = ""&f_delf&"" Then
Set fmi = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&f_member&"'")
Set df = Server.CreateObject("ADODB.RecordSet")
dfSQL = "DELETE * FROM FRIENDS WHERE FRIEND = "&fmi("uye_id")&""
df.open dfSQL,Connection,3,3
Response.Write "<BR><BR><BR>"
Response.Write f_success_del
Response.Write "<BR><BR><BR>"
'df.Close
Set df = Nothing
fmi.Close
Set fmi = Nothing
ElseIF f_action = ""&f_msg&"" Then
Response.Redirect "msgbox.asp?action=new_msg&who="&f_member&""
ElseIF f_action = ""&f_info&"" Then
Set f_m = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&f_member&"'")
Response.Redirect "members.asp?action=member_details&uid="&f_m("uye_id")&""
f_m.Close
Set f_m = Nothing
End IF
End Sub
'###################################################################################################################################
Sub msg_record
mem_type = Request.Form("type")
subject = Request.Form("subject")
msg = Request.Form("message")
If mem_type = "auto" then
mem = Request.Form("mem_names")
ElseIf mem_type = "manual" then
mem = Request.Form("mem_name")
Else
Response.Write "<font class=td_menu>"&msg_error2&"</font>"
End If
Set chk_msg_mem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&mem&"'")
IF chk_msg_mem.Eof Then
Response.Write "<b>" & msg_error3 & "</b>"
ELSE
If subject="" Then
subj = ""&msg_empty&""
Else
subj = subject
End If
IF Request.Form("m_save") = "on" Then
msg_save = True
ELSE
msg_save = False
End IF
If msg="" Then
Response.Write "<font class=td_menu>"&msg_error1&"</font>"
ELSE
Set enter = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM MESSAGES"
enter.open entSQL,Connection,3,3
enter.AddNew
enter("from") = uyebilgi("kul_adi")
enter("to") = mem
enter("msg") = msg
enter("mdate") = Now()
enter("read") = False
enter("subject") = subj
enter("type") = 0
enter("see") = True
enter("msave") = msg_save
enter.Update
enter.Close
Set enter = Nothing
html = html & "<font face=tahoma size=2>Merhaba "&mem&"<br><br>"&Now()&" tarihinde "&uyebilgi("kul_adi")&" isimli üyemiz tarafindan size mesaj yollandi.<BR><BR>Mesaji okumak icin <A HREF="&sett("site_adres")&"/msgbox.asp?action=main><b>BURAYI</b></A> tiklayabilirsiniz.<BR><BR>Bilgilerinize.</font>"
Set Avanos = CreateObject("CDONTS.NewMail")
Avanos.BodyFormat=0
Avanos.MailFormat=0
Avanos.Subject=""&sett("site_isim")&" Sitesinde Yeni Ozel Mesajiniz Var!"
Avanos.Body=HTML
Avanos.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
Avanos.To=""&chk_msg_mem("email")&""
Avanos.Importance = 2
Avanos.Send
set HTML = nothing
set Avanos=nothing
IF msg_save = True Then
success_sent = "<BR><BR><BR>Mesaj Gönderildi ve Kayýt Edildi...<BR><BR><BR><BR>"
ELSE
success_seng = success_sent
End IF
Response.Write "<BR><BR><BR><font class=td_menu><b>"&success_sent&"</b></font><BR><BR><BR><BR>"
END IF
End If
chk_msg_mem.Close
Set chk_msg_mem = Nothing
End Sub
baba_mesaj.Close
Set baba_mesaj = Nothing
gidenmesaj.Close
Set gidenmesaj = Nothing
gelenmesaj.Close
Set gelenmesaj = Nothing
toplamesaj.Close
Set toplamesaj = Nothing
Else
Response.Write "<center><font face=verdana size=2 color="&text&">"&no_entry&"</font></center>"
End If
%>
<!--#include file="inc/adsense.asp" -->
		</td>
	</tr>
</table>
<%
call ORTA2
call ALT
%>