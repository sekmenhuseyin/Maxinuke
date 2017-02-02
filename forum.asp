<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<!--#include file="fckeditor/fckeditor.asp" -->
<%
Set stRs = Connection.Execute("Select * FROM SETTINGS")
	strFloodTime = stRs("flood_time")
	forum_topics = stRs("f_posts")
	forum_replies = stRs("f_replies")
Set stRs = Nothing
IF Session("Enter") = "Yes" Then
Session("Forum_LastTime") = Request.Cookies("MaxiNuke_Forum")("LastTime")
ELSE
Session("Forum_LastTime") = Now()
End IF

act = Request.QueryString("action")
If act = "" Then
call Main
ElseIf act = "Topics" Then
call Topics
ElseIf act = "Topic" Then
call Topic
ElseIf act = "Control" Then
call Control
ElseIf act = "Msg" Then
call msg
ElseIf act = "MsgReg" Then
call reg
ElseIf act = "Rep" Then
call rep
ElseIf act = "RepReg" Then
call repreg
ElseIF act = "Results" Then
call f_results
ElseIF act = "EditPost" Then
call post_edit
ElseIF act = "UpdatePost" Then
call post_update
ElseIF act = "RegQuickReply" Then
call reg_qr
ElseIF act = "siraguncelle" Then
call siraguncelle
End If

Sub reg_qr
IF DateDiff("N" ,Session("FloodTimer"),Now()) > strFloodTime Then
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
amsg = Request.Form("f_rep")
yazan = Session("Member")
IF amsg="" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>Lütfen Mesajýnýzý Yazýnýz...</div><hr color="&tablo_cerceve&">"
elseIf yazan = "" then
response.write "<p class=td>Sayfada cok uzun zaman beklemissiniz ve siteden cikmis gorunuyorsunuz, ya ust taraftan tekrar giris yapin yada alt kisimi kullanarak geri gidin ve yazdiklarinizi kopyalayin.</p>"
ELSE
FOR I = 1 TO UBound(saryHTMLtags)
amsg = Replace(amsg,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1)
Next

Set ms = Connection.Execute("Select * FROM f_mesajlar WHERE mesajid = "&msgid&"")
Set rent = Server.CreateObject("ADODB.RecordSet")
rSQL = "Select * FROM f_mesajlar"
rent.open rSQL,Connection,3,3
rent.AddNew
rent("mmsg") = amsg
rent("yazan") = yazan
rent("tarih") = Now()
rent("okunma") = msgid
rent("kategoriid") = ms("kategoriid")
rent("topic") = False
rent("kilit") = False	
rent.Update
Session("FloodTimerStart") = "Yes"
Set Rmsg = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&Session("Member")&"'")
Session("FloodTimer") = Rmsg("son_tarih")
Set msgUCount = Connection.Execute("UPDATE MEMBERS SET msg_sayisi = "&Rmsg("msg_sayisi")&"+1 WHERE kul_adi='"&Session("Member")&"'")
%>
<!--#include file="forum_abone.asp" -->
<%
Response.Redirect("?action=Topic&id=" & msgid)
END IF
Else
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>" class="td_menu">
	<tr>
		<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B>HATA</B></CENTER></td>
	</tr>
<SCRIPT>
function countDown()
{
var newValue = parseInt(document.geneyaz.numberShown.value)-1
document.geneyaz.numberShown.value = newValue
myTimer=setTimeout("countDown()",1000)
}
</SCRIPT>
<FORM name=geneyaz>
	<tr>
		<td align="center" valign="top"><BR><BR>Arka Arkaya Kisa Sürede Cevap Veremezsiniz !<BR><BR>Tekrar yazabilmeniz icin <INPUT size=1 class="form1" name=numberShown size="1">
		<SCRIPT language=JavaScript>
		document.geneyaz.numberShown.value="60"
		countDown()
		</SCRIPT>
		Saniye Beklemelisiniz.<BR><BR>
<!-- 		<a href="#"><button class="form1" onclick="parentElement.click();">Tekrar Gonder</button></a><BR><BR><BR> -->
		</td>
	</tr>
</FORM>
</table>
<%
End IF
End Sub

Sub post_update
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
amsg = Request.Form("amsg")
sc = Request.Form("f_ep_sc")

IF amsg="" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&f_error2&"</div><hr color="&tablo_cerceve&">"
ElseIF sc = "" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&sc_err1&"</div><hr color="&tablo_cerceve&">"
ElseIF Session("F_EP_SC") <> sc Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&sc_err2&"</div><hr color="&tablo_cerceve&">"
ELSE
FOR I = 1 TO UBound(saryHTMLtags)
amsg = Replace(amsg,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1)
Next

Set rent = Server.CreateObject("ADODB.RecordSet")
rSQL = "Select * FROM f_mesajlar WHERE mesajid = "&msgid&""
rent.open rSQL,Connection,3,3

IF rent("yazan") = Session("Member") Then

rent("mmsg") = amsg
rent("tarih") = Now()
rent.Update

IF rent("topic") = True Then
msgid = msgid
ELSE
msgid = rent("okunma")
End IF

Response.Redirect "?action=Topic&id="&msgid&""
End IF
END IF
End Sub

Sub post_edit
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM f_mesajlar WHERE mesajid = "&QS_CLEAR(Request.QueryString("id"))&"",Connection,3,3
IF rs("yazan") = Session("Member") Then
Session("F_EP_SC") = CPASS(13)
%>
<form action="?action=UpdatePost&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<table border="0" cellspacing="5" cellpadding="3" width="80%" align="center" class="td_menu" style="font-weight:bold" bgcolor="<%=background%>">
	<tr>
		<td bgcolor="<%=content_bg%>"><%=security_code%> : <%=Session("F_EP_SC")%><input type="text" name="f_ep_sc" size="4" class="form1"></td>
	</tr>
	<tr>
		<td bgcolor="<%=content_bg%>" colspan="2">
		<%
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.BasePath = "FCKeditor/"
		oFCKeditor.Value	= rs("mmsg")
		oFCKeditor.Create "amsg"
		%>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="<%=background%>"><input type="submit" class="submit" value="<%=uye_kayit%>" style="width:100%"></td>
	</tr>
</table>
</form>
<%
ELSE
Response.Write "<hr color="&tablo_cerceve&"><div align=center>Baþkasýna ait konuyu düzenleyemezsiniz !</div><hr color="&tablo_cerceve&">"
End IF
rs.Close
Set rs = Nothing
End Sub

Sub siraguncelle
Set siraguncell=Server.CreateObject("Adodb.Recordset")
siraguncell.ActiveConnection = Connection
Sorgu="SELECT * FROM f_anakategori WHERE anakatid="&Request.QueryString("anakatid")&""
siraguncell.open Sorgu,Connection,1,3
	If Request.QueryString("nere")="ust" Then
	sirasi=siraguncell("sira")-1
	elseIf Request.QueryString("nere")="alt" Then
	sirasi=siraguncell("sira")+1
	End if
siraguncell.update "sira",sirasi
siraguncell.close
set siraguncell=nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End sub

Sub f_results
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM f_mesajlar WHERE kategoriid = "&Request.Form("forums")&" AND baslik LIKE '%"&duz(Request.Form("word"))&"%' AND topic = True",Connection,3,3
%></b>
<table border="0" cellspacing="1" cellpadding="1" width="100%" class="td_menu" bgcolor="<%=content_bg%>" align="center">
<tr height="20">
<td colspan="2" width="50%" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%=nm_title%></td>
<td width="20%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%=fr_pop5%></td>
<td width="30%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%=fw_pop5%></td>
</tr>
<%
Do Until rs.EoF
Set replies = Connection.Execute("Select Count(*) AS count FROM f_mesajlar WHERE okunma = "&rs("mesajid")&" AND topic = False")
Set memInfo = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs("yazan")&"'")
%>
<tr bgcolor="<%=content_bg%>">
<td width="1%">
<%
IF rs("kilit") = True then %>
<img src="images/temalar/<%=Session("SiteTheme")%>/locked_msg.gif">
<% ElseIF replies("count") >= 10 Then %>
<img src="images/temalar/<%=Session("SiteTheme")%>/hot_msg.gif">
<% Else %>
<img src="images/temalar/<%=Session("SiteTheme")%>/normal_msg.gif">
<% End If %>
</td>
<td width="50%"><a href="?action=Topic&id=<%=rs("mesajid")%>"><%=rs("baslik")%></a></td>
<td width="20%" align="center"><%=replies("count")%></td>
<td width="30%" align="center">
<%
IF memInfo.EoF OR memInfo.BoF Then
Response.Write rs("yazan")
ELSE
Response.Write "<a href=""Members.asp?action=member_details&uid="&memInfo("uye_id")&""">" & rs("yazan") & "</a>"
End IF
%>
</td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
With Response
	.Write "<br>"
	.Write "<div align=center class=td_menu>"
	.Write fs_result_1 & "<b>" & rs.RecordCount & "</b>" & fs_result_2
	.Write "</div>"
End With
rs.Close
Set rs = Nothing
End Sub

Sub Main
Set mcats = Connection.Execute("Select * FROM f_anakategori ORDER BY sira")
Do While NOT mcats.Eof
Set cats = Connection.Execute("Select * FROM f_kategoriler WHERE anakatid = "&mcats("anakatid")&" ORDER BY katad")
%>
<script language="JavaScript">
function anakatsill(listForm2)
{ 
   listForm2.target="_self"; 
   listForm2.action="";
   var answer = confirm ("Ana Kategorinizi Silmek Istediðinize Emin misiniz? Bu Ana Kategoriyi Sildiginiz Zaman buna bagli tum alt forumlar ve icerisindeki mesajlar geri donusumu olmaksizin silinecektir.") 
   if (answer)
       return true;
   else
       return false;
} 
//-->
</script>
<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="<%=content_bg%>" align="center" class="td_menu">
	<tr style="font-weight:bold">
		<td width="2%" height="20" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%IF Session("Level")="1" OR Session("Level")="5" Then%><A HREF="?action=siraguncelle&nere=ust&anakatid=<%=mcats("anakatid")%>&sira=<%=mcats("sira")%>"><IMG SRC="images/yukari.png" BORDER="0" ALT="Yukarý Taþý"></A><A HREF="?action=siraguncelle&nere=alt&anakatid=<%=mcats("anakatid")%>&sira=<%=mcats("sira")%>"><IMG SRC="images/assa.png" BORDER="0" ALT="Aþþaðý Taþý"></A><%End if%></td>
		<td width="58%" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">&nbsp;<%=mcats("anakatadi")%><% IF Session("Level")="1" OR Session("Level")="5" Then%> (<a href="forum_a.asp?action=anaduzenle&id=<%=mcats("anakatid")%>">Düzenle</a> - <a onClick="return anakatsill(this)" href="forum_a.asp?action=anakatsil&id=<%=mcats("anakatid")%>">Sil</a>)<%End IF %></div></td>
		<td width="30%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">Son Mesaj</td>
		<td width="5%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">Konular</td>
		<td width="5%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">Mesajlar</td>
	</tr>
<%
Do While NOT cats.Eof
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = forum_bg_1
Case Else
color = forum_bg_2
End Select

Set tcount = Connection.Execute("Select Count(*) AS count FROM f_mesajlar WHERE kategoriid = "&cats("kategoriid")&" AND topic = True")
Set acount = Connection.Execute("Select Count(*) AS count FROM f_mesajlar WHERE kategoriid = "&cats("kategoriid")&" AND topic = False")
Set last = Connection.Execute("Select * FROM f_mesajlar WHERE kategoriid = "&cats("kategoriid")&" ORDER BY tarih DESC")
If last.Eof Then
last_post = "-"
last_pmember = "-"
last_pmesajid = ""
ELSE
last_post = last("tarih")
okunma = last("okunma")
sonakategoriid = last("kategoriid")
Set last3 = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&last("yazan")&"'")
IF last3.Eof OR last3.Bof Then
last_pmember = "###"
last_pmesajid = "###"
ELSE
last_pmember = last3("kul_adi")
last_pmesajid = last3("uye_id")
End IF
End If
%>

<tr class="td_menu">

<td width="2%" bgcolor="<%=color%>">
<%
IF ""&last_post&"" > ""&Session("Forum_LastTime")&"" Then
FLT = True
ELSE
FLT = False
End IF
IF cats("kilit") = True then
Response.Write "<img src=images/temalar/"&Session("SiteTheme")&"/locked.gif>"
ElseIF cats("noentry") = True Then
Response.Write "<img src=images/temalar/"&Session("SiteTheme")&"/no_entry_forum.gif>"
ElseIF FLT = True Then
Response.Write "<img src=images/temalar/"&Session("SiteTheme")&"/forum_open_new.gif>"
ELSE
Response.Write "<img src=images/temalar/"&Session("SiteTheme")&"/forum_open.gif>"
End If
%>
</td>




<script language="JavaScript">
function katsill(listForm2)
{ 
   listForm2.target="_self"; 
   listForm2.action="";
   var answer = confirm ("Kategorinizi Silmek Istediðinize Emin misiniz? Bu Ana Kategoriyi Sildiginiz Zaman icerisindeki mesajlar geridonusumu olmaksizin silinecektir.") 
   if (answer)
       return true;
   else
       return false;
} 
//-->
</script>
<td bgcolor="<%=color%>" width="<%=genislik%>"><b><a href="?action=Control&id=<%=cats("kategoriid")%>"><%=cats("katad")%></a></b><% If Session("Level") = "1" OR Session("Level") = "5" Then %> (<a href="forum_a.asp?action=EditCat&id=<%=cats("kategoriid")%>"><%=f_edit%></a> - <a onClick="return katsill(this)" href="forum_a.asp?action=DeleteCat&id=<%=cats("kategoriid")%>"><%=f_del%></a> - <% If cats("kilit") = True Then %><a href="forum_a.asp?action=UnLock&id=<%=cats("kategoriid")%>"><%=f_unlock%></a><%ElseIf cats("kilit") = False Then%><a href="forum_a.asp?action=Lock&id=<%=cats("kategoriid")%>"><%=f_lock%></a><% End If %> - <% If cats("noentry") = True Then %><a href="forum_a.asp?action=SetEntry&id=<%=cats("kategoriid")%>"><%=f_entry%></a><%ElseIf cats("noentry") = False Then%><a href="forum_a.asp?action=SetNoEntry&id=<%=cats("kategoriid")%>"><%=f_noentry%></a><% End If %>)<% End IF %><br><%=cats("kataciklama")%></td>
<td bgcolor="<%=color%>">
<%
'Konuda yok ise
If tcount("count")=0 Then
response.write"Bu ana kategori altýnda henüz hiç konu açýlmamýþ."
'1 yada daha fazla konu varsa
ELSEIF tcount("count")>0 Then


'son yazilan mesaji bul
set sonyazilan = server.createobject("adodb.recordset")
sonyazilansql = "select * from f_mesajlar where topic=true and kategoriid = "&sonakategoriid&" order by mesajid desc"
sonyazilan.open sonyazilansql, Connection, 1, 3
%>
<B><A HREF="?action=Topic&id=<%=sonyazilan("mesajid")%>"><%=sonyazilan("baslik")%></A></B>
<br>
Yazan : <a href="members.asp?action=member_details&uid=<%=last_pmesajid%>"><%=last_pmember%></a>
<BR>
<%=last_post%>
<%
sonyazilan.close
set sonyazilan=nothing

End IF
%>
</td>
<td align="center" bgcolor="<%=content_bg%>"><%=tcount("count")%></td>
<td align="center"bgcolor="<%=color%>"><%=acount("count")%></td>
</tr>

<%
cats.MoveNext
Loop
%>
</table><br>
<%
mcats.MoveNext
Loop
%>
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu" bgcolor="<%=content_bg%>">
	<tr>
		<td align="center"><img src="images/temalar/<%=Session("SiteTheme")%>/forum_open.gif" alt="<%=bf_normal%>" align="absmiddle">&nbsp;<%=bf_normal%> - <img src="images/temalar/<%=Session("SiteTheme")%>/forum_open_new.gif" alt="<%=bf_new%>" align="absmiddle">&nbsp;<%=bf_new%> - <img src="images/temalar/<%=Session("SiteTheme")%>/locked.gif" alt="<%=bf_locked%>" align="absmiddle">&nbsp;<%=bf_locked%> - <img src="images/temalar/<%=Session("SiteTheme")%>/no_entry_forum.gif" alt="<%=bf_noentry%>" align="absmiddle">&nbsp;<%=bf_noentry%></td>
	</tr>
</table>
<%
End Sub
Sub Topics
If Application("Per") = "OK" Then
id = Request.QueryString("id")
Set msgs = Server.CreateObject("ADODB.RecordSet")
msgs.open "Select * FROM f_mesajlar WHERE kategoriid="&id&" AND topic=True ORDER BY tarih desc ",Connection,3,3
Set cat = Connection.Execute("Select * FROM f_kategoriler WHERE kategoriid = "&id&"")

Page = Request.QueryString("Page")
IF Page="" Then
Page = "1"
End IF
%>
</b>
<%
If Session("Enter") = "Yes" Then
%>
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" bgcolor="<%=content_bg%>" class="td_menu">
 <TR>
	<TD> <B><A HREF="forum.asp"><%=sett("site_isim")%> Forum </A>» <%=cat("katad")%></B></TD>
	<TD align="right">
	<%
	If cat("kilit") = True Then
	Response.Write "<b>Bu Forum Kilitlenmiþtir Yeni Konu açamazsýnýz.</b>"
	Else
	%>
	<a href="?action=Msg&id=<%=id%>"><img src="images/temalar/<%=Session("SiteTheme")%>/f_newpost.gif" border="0"></a>
	<%End if%>	
	</TD>
 </TR>
 </TABLE>
<% End If
IF msgs.Eof Then
Response.Write "<center><font face=verdana size=2>"&no_topics&"</center>"
ELSE %>
<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="<%=content_bg%>" align="center" class="td_menu">
<tr bgcolor="<%=content_bg%>" class="td_menu" style="font-weight:bold" height="20">
<td width="2%" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">&nbsp;</td>
<td width="76%" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">&nbsp;Konu / Konuyu Baþlatan </td>
<td width="12%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%=m_lastpost%></td>
<td width="5%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif">Cevaplar</td>
<td width="5%" align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif"><%=m_read%></td>

</tr>
<%
msgs.pagesize = forum_topics
msgs.absolutepage = Page
sayfa = msgs.pagecount
For I = 1 To msgs.pagesize
IF msgs.eof Then Exit For

number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = ""&forum_bg_1&""
Case Else
color = ""&forum_bg_2&""
End Select

Set acount = Connection.Execute("Select Count(*) AS count FROM f_mesajlar WHERE okunma = "&msgs("mesajid")&" AND topic = False")
Set last = Connection.Execute("Select * FROM f_mesajlar WHERE mesajid = "&msgs("mesajid")&" OR okunma = "&msgs("mesajid")&" ORDER BY tarih DESC")
If last.Eof Then
last_post = "###"
last_pmember = "###"
ELSE
last_post = last("tarih")
Set last3 = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&last("yazan")&"'")
IF last3.Eof OR last3.Bof Then
last_pmember = "###"
last_pmemesajid = ""
ELSE
last_pmember = last3("kul_adi")
last_pmemesajid = last3("uye_id")
End IF
End If

Set wMem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&msgs("yazan")&"'")
IF wMem.Eof OR wMem.Bof Then
memberid = ""
ELSE
memberid = wMem("uye_id")
End IF
%>
<tr>
<td bgcolor="<%=color%>">
<%
If msgs("kilit") = True then %>
<img src=images/temalar/<%=Session("SiteTheme")%>/locked_msg.gif>
<% ElseIf acount("count") >= 10 Then %>
<img src="images/temalar/<%=Session("SiteTheme")%>/hot_msg.gif">
<% Else %>
<img src="images/temalar/<%=Session("SiteTheme")%>/normal_msg.gif">
<% End If %>
</td>
<td bgcolor="<%=color%>"><b><a href="?action=Topic&id=<%=msgs("mesajid")%>"><%=duz(msgs("baslik"))%></a></b><% If Session("Level") = "1" OR Session("Level") = "5" Then %> (<a href="?action=EditPost&id=<%=msgs("mesajid")%>"><%=f_edit%></a> - <a href="forum_a.asp?action=DeleteMsg&id=<%=msgs("mesajid")%>"><%=f_del%></a> - <% If msgs("kilit") = False Then %><a href="forum_a.asp?action=LockMsg&id=<%=msgs("mesajid")%>"><%=f_lock%></a><% Else %><a href="forum_a.asp?action=UnLockMsg&id=<%=msgs("mesajid")%>"><%=f_unlock%></a><% End If %>)<% End IF %><BR><a href="members.asp?action=member_details&uid=<%=memberid%>"><%=duz(msgs("yazan"))%></a></td>
<td bgcolor="<%=color%>"><%=last_post%><br>Yazan : <a href="members.asp?action=member_details&uid=<%=last_pmemesajid%>"><%=last_pmember%></a></td>
<td align="center" bgcolor="<%=color%>"><%=acount("count")%></td>
<td align="center" bgcolor="<%=color%>"><%=msgs("okunma")%></td>

</tr>
<%
msgs.MoveNext
Next
%>
</table>
<br>
<div align="right" class="td_menu">
  <%
bp = Page-1
IF Page <> 1 Then
With Response
	.Write "<a href=""?action=Topics&id="&id&"&Page="&bp&""">« "&previous_page&"</a>"
	.Write "&nbsp;-&nbsp;"
End With
End IF
for y=1 to sayfa
if s=y then
response.write y
else
response.write "<a href=""?action=Topics&id="&id&"&Page="&y&"""><font class=""td_menu"">["&y&"]</a>"
end if
next
np = Page+1
IF NOT msgs.Eof Then
With Response
	.Write "&nbsp;-&nbsp;"
	.Write "<a href=""?action=Topics&id="&id&"&Page="&np&""">"&next_page&" »</a>"
End With
End IF
%>
</div><br>
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu" bgcolor="<%=content_bg%>">
	<tr>
		<td align="center"><img src="images/temalar/<%=Session("SiteTheme")%>/normal_msg.gif" alt="<%=bm_normal%>">&nbsp;<%=bm_normal%> - <img src="images/temalar/<%=Session("SiteTheme")%>/hot_msg.gif" alt="<%=bm_hot%>">&nbsp;<%=bm_hot%> - <img src="images/temalar/<%=Session("SiteTheme")%>/locked_msg.gif" alt="<%=bm_locked%>">&nbsp;<%=bm_locked%></td>
	</tr>
</table>


<%
End If
End If
End Sub
Sub Topic

tid = Request.QueryString("id")
tid = QS_CLEAR(tid)
Set rs = Server.CreateObject("ADODB.RecordSet")
rSQL = "Select * FROM f_mesajlar WHERE mesajid = "&tid&" AND topic = True"
rs.open rSQL,Connection,3,3
Set checkCat = Connection.Execute("Select * FROM f_kategoriler WHERE kategoriid = "&rs("kategoriid")&"")

If checkCat("kilit") = True Then
Response.Write "<center><font face=arial size=3><b>"&f_locked&"</b></center>"
ELSE

Page = Request.QueryString("Page")
Page = QS_CLEAR(Page)
IF Page="" Then
Page = "1"
End IF

rs("okunma") = rs("okunma") + 1
rs.Update

Set frep = Server.CreateObject("ADODB.RecordSet")
frep.open "Select * FROM f_mesajlar WHERE okunma = "&tid&" AND topic = False ORDER BY mesajid",Connection,3,3
Set cat = Connection.Execute("Select * FROM f_kategoriler WHERE kategoriid = "&rs("kategoriid")&"")
Set mem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs("yazan")&"'")

IF mem.Eof Then
f_mem_name = "#########"
f_mem_msg = ""
f_mem_age = ""
f_mem_signature = ""
f_mem_id = ""
f_mem_level = "0"
f_mem_online = "?"
resim="yok.gif"
ELSE
f_mem_name = mem("kul_adi")
f_mem_msg = mem("msg_sayisi")
strAge = Int(Right(mem("yas"),4))
strNow = Int(Right(Date(),4))
f_mem_age = strNow - strAge
f_mem_signature = mem("imza")
f_mem_id = mem("uye_id")
f_mem_level = mem("seviye")
	If mem("resmim_onay")=true Then
	resim=mem("resmim")
	Else
	resim="yok.gif"
	End if
	IF mem("u_online") = True Then
	f_mem_online = member_online
	ELSE
	f_mem_online = member_offline
	End IF
End IF
%>
<!--#include file="inc/forum_ratings.asp" -->
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" bgcolor="<%=content_bg%>" class="td_menu">
 <TR>
	<TD> <B><A HREF="forum.asp"><%=sett("site_isim")%> Forum </A>» <a href="?action=Topics&id=<%=cat("kategoriid")%>"><%=cat("katad")%></a> » <%=SmiLey(duz(rs("baslik")))%></B></TD>
	<TD align="right">
	<% If rs("kilit") = True Then
	Response.Write "<center><font face=tahoma style=font-size:11px;><b>"&m_locked&"</b></center>"
	Else
	%>
	<a href="?action=Rep&id=<%=tid%>" method="post"><img src="images/temalar/<%=Session("SiteTheme")%>/f_newreply.gif" border="0"></a>
	<% End If %>
	</TD>
 </TR>
 </TABLE>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" class="td_menu">
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/yazi.gif" align="left">&nbsp;<%=rs("tarih")%></TD>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" align="right">
		<%
		bp = Page-1
		IF Page <> 1 Then
		With Response
			.Write "<a href=""?action=Topic&id="&QS_CLEAR(Request.QueryString("id"))&"&Page="&bp&""">« "&previous_page&"</a>"
			.Write "&nbsp;-&nbsp;"
		End With
		End IF
		for y=1 to sayfa
		if s=y then
		response.write y
		else
		response.write "<a href=""?action=Topic&id="&QS_CLEAR(Request.QueryString("id"))&"&Page="&y&"""><font class=""td_menu"">["&y&"]</a>"
		end if
		next
		np = Page+1
		IF NOT frep.Eof Then
		With Response
			.Write "&nbsp;-&nbsp;"
			.Write "<a href=""?action=Topic&id="&QS_CLEAR(Request.QueryString("id"))&"&Page="&np&""">"&next_page&" »</a>"
		End With
		End IF
		%>		
		</TD>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" align="right"><%IF f_mem_name = Session("Member") Then%><a href="?action=EditPost&id=<%=rs("mesajid")%>"><B>Duzenle</B></a><%End IF%> - <a href="?action=Rep&Type=Quote&id=<%=Request.QueryString("id")%>&qid=<%=rs("mesajid")%>"><img src="images/temalar/<%=Session("SiteTheme")%>/forum_quote.gif" border="0"></a></TD>
	</TR>
 </TABLE>
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
	<TR bgcolor="<%=content_bg%>">
		<TD valign="top" width="15%">
			<table border="0" cellspacing="1" cellpadding="5" align="center" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
			<TR bgcolor="<%=content_bg%>">
				<TD rowspan="5" align="center" width="60"><a href="members.asp?action=member_details&uid=<%=f_mem_id%>"><IMG SRC="uploads/avatar/<%=resim%>" alt="<%=f_mem_name%>" WIDTH="120" BORDER="0"></A></TD>
				<TD align="center"><b><a href="members.asp?action=member_details&uid=<%=f_mem_id%>"><%=f_mem_name%></a></b></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center"><%=fm_level%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center">Mesaj : <%=f_mem_msg%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center">Yaþ : <%=f_mem_age%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center"><%=fm_include_img2%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center" colspan="2"><%=fm_include_img%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center" colspan="2"><IMG SRC="images/temalar/b/user_<%=f_mem_online%>.gif"></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center" colspan="2"><a href="msgbox.asp?action=new_msg&who=<%=f_mem_name%>">PM At</a> - <a href="members.asp?action=add_friendlist&id=<%=f_mem_id%>">Listeme Ekle</a></TD>
			</TR>
			</TABLE>
<!--#include file="inc/adsense_2.asp" -->
		</TD>
		<TD valign="top" width="85%">
		<B><%=SmiLey(rs("baslik"))%></B>
		<Hr color="<%=tablo_cerceve%>" size="1">
		<%Response.Write SmiLey(rs("mmsg"))%>
		<Hr color="<%=tablo_cerceve%>" size="1">
		<%=f_mem_signature%>
<!--#include file="inc/adsense_3.asp" -->
		</TD>
	</TR>
 </TABLE>
<%
mem.Close
Set mem = Nothing
IF frep.Eof OR frep.Bof Then
Response.Write "<div align=center class=td_menu><b>"&f_error3&"</b></div>"
ELSE
frep.pagesize = forum_replies
frep.absolutepage = Page
sayfa = frep.pagecount
For I = 1 To frep.pagesize
IF frep.eof Then Exit For

Set mem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&frep("yazan")&"'")

IF mem.EoF Then
f_mem_name = "###"
f_mem_msg = ""
f_mem_age = ""
f_mem_signature = ""
f_mem_id = ""
f_mem_level = "0"
f_mem_online = "?"
resim="yok.gif"
ELSE
f_mem_name = mem("kul_adi")
f_mem_msg = mem("msg_sayisi")
strAge = Int(Right(mem("yas"),4))
strNow = Int(Right(Date(),4))
f_mem_age = strNow - strAge
f_mem_signature = mem("imza")
f_mem_id = mem("uye_id")
f_mem_level = mem("seviye")
	If mem("resmim_onay")=true Then
	resim=mem("resmim")
	Else
	resim="yok.gif"
	End if
	IF mem("u_online") = True Then
	f_mem_online = member_online
	ELSE
	f_mem_online = member_offline
	End IF
End IF
%>
<!--#include file="inc/forum_ratings.asp" -->
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" class="td_menu">
	<TR>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/yazi.gif" align="left">&nbsp;<%=frep("tarih")%></TD>
		<TD background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" align="right"><% IF Session("Level") = "1" OR Session("Level") = "5" Then %><a href="forum_a.asp?action=Edit-Rep&id=<%=frep("mesajid")%>">Düzenle</a> <b>|</b> <a href="forum_a.asp?action=Del-Rep&id=<%=frep("mesajid")%>">Sil</a> - <% End IF %><%IF f_mem_name = Session("Member") Then%><a href="?action=EditPost&id=<%=frep("mesajid")%>">Duzenle - <% End IF %><a href="?action=Rep&Type=Quote&id=<%=Request.QueryString("id")%>&qid=<%=frep("mesajid")%>"><img src="images/temalar/<%=Session("SiteTheme")%>/forum_quote.gif" border="0"></a></TD>
	</TR>
 </TABLE>
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
	<TR bgcolor="<%=content_bg%>">
		<TD valign="top" width="15%">			
			<table border="0" cellspacing="1" cellpadding="5" align="center" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
			<TR bgcolor="<%=content_bg%>">
				<TD rowspan="5" align="center" width="60"><a href="members.asp?action=member_details&uid=<%=f_mem_id%>"><IMG SRC="uploads/avatar/<%=resim%>" alt="<%=f_mem_name%>" WIDTH="120" BORDER="0"></A></TD>
				<TD align="center"><b><a href="members.asp?action=member_details&uid=<%=f_mem_id%>"><%=f_mem_name%></a></b></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center"><%=fm_level%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center">Mesaj : <%=f_mem_msg%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center">Yaþ : <%=f_mem_age%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center"><%=fm_include_img2%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center" colspan="2"><%=fm_include_img%></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center" colspan="2"><IMG SRC="images/temalar/b/user_<%=f_mem_online%>.gif"></TD>
			</TR>
			<TR bgcolor="<%=content_bg%>">
				<TD align="center" colspan="2"><a href="msgbox.asp?action=new_msg&who=<%=f_mem_name%>">PM At</a> - <a href="members.asp?action=add_friendlist&id=<%=f_mem_id%>">Listeme Ekle</a></TD>
			</TR>
			</TABLE>
<!--#include file="inc/adsense_2.asp" -->
		</TD>
		<TD valign="top" width="85%">
		<%
		strReplyMsg = SmiLey(frep("mmsg"))
		IF InStr(strReplyMsg,"[Quote:") Then
		strReplyMsg = Quote(strReplyMsg)
		ELSE
		strReplyMsg = strReplyMsg
		End IF
		Response.Write strReplyMsg
		%>
		<Hr color="<%=tablo_cerceve%>" size="1">
		<%=f_mem_signature%>
		</TD>
	</TR>
 </TABLE>
<%
frep.MoveNext
Next
End IF
%>
<table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="<%=tablo_cerceve%>" class="td_menu">
 <TR>
	<TD align="right"><B>
	<%
	bp = Page-1
	IF Page <> 1 Then
	With Response
		.Write "<a href=""?action=Topic&id="&QS_CLEAR(Request.QueryString("id"))&"&Page="&bp&""">« "&previous_page&"</a>"
		.Write "&nbsp;-&nbsp;"
	End With
	End IF
	for y=1 to sayfa
	if s=y then
	response.write y
	else
	response.write "<a href=""?action=Topic&id="&QS_CLEAR(Request.QueryString("id"))&"&Page="&y&"""><font class=""td_menu"">["&y&"]</a>"
	end if
	next
	np = Page+1
	IF NOT frep.Eof Then
	With Response
		.Write "&nbsp;-&nbsp;"
		.Write "<a href=""?action=Topic&id="&QS_CLEAR(Request.QueryString("id"))&"&Page="&np&""">"&next_page&" »</a>"
	End With
	End IF
	%></B>
	</TD>
 </TR>
 </TABLE>





<%
IF Session("Enter") = "Yes" Then
	If rs("kilit") = True Then
	Response.Write "<center><font face=tahoma style=font-size:11px;><b>"&m_locked&"</b></center>"
	Else
	%>
<table border="0" cellspacing="0" cellpadding="0" width="50%" class="td_menu" align="center">
<form action="?action=RegQuickReply&id=<%=Request.QueryString("id")%>" method="post">
	<tr>
		<td align="center"><b>HIZLI CEVAP</b></td>
	</tr>
	<tr>
		<td align="center"><textarea name="f_rep" rows="10" style="width:100%" class="form1"></textarea></td>
	</tr>
	<tr>
		<td align="center"><input type="submit" value="Hýzlý Cevabý Gönder" class="submit"><BR><BR><a href="?action=Rep&id=<%=tid%>"><button class=submit onclick="parentElement.click();">Geliþmiþ Cevaba Git</button></a></td>
	</tr>
</form>
</table>
<%
End If
End IF
END IF
End Sub
Sub msg
kategoriid = Request.QueryString("id")
kategoriid = QS_CLEAR(kategoriid)
IF Session("Enter") = "Yes" Then
Set cat = Connection.Execute("Select * FROM f_kategoriler WHERE kategoriid = "&QS_CLEAR(Request.QueryString("id"))&"")
Response.Write "» " & cat("katad")
IF cat("kilit") = True Then
Response.Write "<hr size=1 color="&tablo_cerceve&"><div align=center>" & f_locked & "</div><hr size=1 color="&tablo_cerceve&">"
ELSE
f1_sec_code = CPASS(13)
Session("Forum1SecurityCode") = f1_sec_code
%>
<form action="?action=MsgReg&id=<%=kategoriid%>" method="post">
<table border="0" cellspacing="3" cellpadding="3" width="80%" align="center" class="td_menu" style="font-weight:bold" bgcolor="<%=background%>">
	<tr>
		<td bgcolor="<%=content_bg%>"><%=nm_title%> : <input type="text" name="baslik" class="form1" size="70"></td>
		<td bgcolor="<%=content_bg%>"><%=security_code%> : <%=Session("Forum1SecurityCode")%><input type="text" name="f1_sec_code" size="4" class="form1"></td>
	</tr>
	<tr>
		<td bgcolor="<%=content_bg%>" colspan="2">
		<%
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.BasePath = "FCKeditor/"
		oFCKeditor.Value	= ""
		oFCKeditor.Create "mmsg"
		%>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="<%=background%>"><input type="submit" class="submit" value="<%=uye_kayit%>" style="width:100%" onClick="this.form.submit();this.disabled=true; return true;"></td>
	</tr>
</table>
</form>
<%
End IF
Else
Response.Write WriteMsg(no_entry)
End IF
End Sub 

Sub reg
IF DateDiff("N" ,Session("FloodTimer"),Now()) > strFloodTime Then
kategoriid = Request.QueryString("id")
kategoriid = QS_CLEAR(kategoriid)
title = Request.Form("baslik")
mmsg = Request.Form("mmsg")
scode = duz(Request.Form("f1_sec_code"))
yazan = Session("Member")
If title="" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu><b>"&f_error1&"</b></div><hr color="&tablo_cerceve&">"
ElseIf mmsg="" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu><b>"&f_error2&"</b></div><hr color="&tablo_cerceve&">"
ElseIF scode = "" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&sc_err1&"<b></b></div><hr color="&tablo_cerceve&">"
ElseIF scode <> Session("Forum1SecurityCode") Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&sc_err2&"<b></b></div><hr color="&tablo_cerceve&">"
elseIf yazan = "" then
response.write "<p class=td>Sayfada cok uzun zaman beklemissiniz ve siteden cikmis gorunuyorsunuz, ya ust taraftan tekrar giris yapin yada alt kisimi kullanarak geri gidin ve yazdiklarinizi kopyalayin.</p>"
ELSE
FOR I = 1 TO UBound(saryHTMLtags)
mmsg = Replace(mmsg,"<"&saryHTMLtags(""&I&"")&">","{"&saryHTMLtags(""&I&"")&"}",1,-1,1)
Next

Set fent = Server.CreateObject("ADODB.RecordSet")
fSQL = "Select * FROM f_mesajlar"
fent.open fSQL,Connection,3,3
fent.AddNew
fent("baslik") = title
fent("mmsg") = mmsg
fent("yazan") = yazan
fent("tarih") = Now()
fent("okunma") = 0
fent("kategoriid") = kategoriid
fent("topic") = True
fent.Update
Session("FloodTimerStart") = "Yes"
Set Cmsg = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&Session("Member")&"'")
Session("FloodTimer") = Cmsg("son_tarih")
Set msgCount = Connection.Execute("UPDATE MEMBERS SET msg_sayisi = "&Cmsg("msg_sayisi")&"+1 WHERE kul_adi='"&Session("Member")&"'")
html = "<font face=verdana size=2>"&sett("site_isim")&" Forumlarýna "&Session("Member")&" tarafindan yeni bir konu eklendi. Kategoriye Gitmek icin <A HREF="&sett("site_adres")&"/forum.asp?action=Topics&id="&kategoriid&">BURAYA</A> tiklayabilirsiniz<BR><BR>Bilgilerinize.</font>"
Set mail = CreateObject("CDONTS.NewMail")
mail.BodyFormat=0
mail.MailFormat=0
mail.Subject=""&sett("site_isim")&" Foruma Konu Eklendi"
mail.Body=HTML
mail.From= ""&sett("site_isim")&"<"&sett("site_mail")&">"
mail.to=sett("site_mail")
mail.Importance=1
mail.Send
set HTML = nothing
set mail=nothing
Response.Redirect "?action=Topics&id="&kategoriid&""
END IF
ELSE
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu><b>"&f_error5&"</b></div><hr color="&tablo_cerceve&">"
End IF
End Sub


Sub rep
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
Set rmsg = Connection.Execute("Select * FROM f_mesajlar WHERE mesajid = "&msgid&"")
If rmsg("kilit") = True Then
Response.Write "<hr size=1 color="&tablo_cerceve&"><div align=center><b>"&m_locked&"</b></div><hr size=1 color="&tablo_cerceve&">"
ELSE
IF Session("Enter") = "Yes" Then
f2_sec_code = CPASS(13)
Session("Forum2SecurityCode") = f2_sec_code

IF Request.QueryString("qid") <> "" Then
strRequestID = Request.QueryString("qid")
ELSE
strRequestID = Request.QueryString("id")
End IF

Set iqmRs = Server.CreateObject("ADODB.RecordSet")
iqmRs.Open "Select * FROM f_mesajlar WHERE mesajid = " & strRequestID,Connection,3,3

IF Request.QueryString("Type") = "Quote" Then
strReplyMsg = "[Quote:" & iqmRs("yazan") & "]" & iqmRs("mmsg") & "[/Quote]"
ELSE
strReplyMsg = ""
End IF

iqmRs.Close
Set iqmRs = Nothing
%>
<form action="?action=RepReg&id=<%=msgid%>" method="post">
<% IF Request.QueryString("Type") = "Quote" Then %>
<input type="hidden" name="quote" value="Yes">
<% End IF %>
<table border="0" cellspacing="5" cellpadding="3" width="80%" align="center" class="td_menu" style="font-weight:bold" bgcolor="<%=background%>">
	<tr>
		<td bgcolor="<%=content_bg%>"><%=security_code%> : <%=Session("Forum2SecurityCode")%><input type="text" name="f2_sec_code" size="4" class="form1"></td>
		<td bgcolor="<%=content_bg%>">Bu konuya abone olmak istiyorum. <INPUT TYPE="checkbox" NAME="abone" value="ol"></td>
	</tr>
	<tr>
		<td bgcolor="<%=content_bg%>" colspan="2">
		<%
		Dim oFCKeditor
		Set oFCKeditor = New FCKeditor
		oFCKeditor.BasePath = "FCKeditor/"
		oFCKeditor.Value	= strReplyMsg
		oFCKeditor.Create "amsg"
		%>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="<%=background%>"><input type="submit" class="submit" value="<%=uye_kayit%>" style="width:100%" onClick="this.form.submit();this.disabled=true; return true;"></td>
	</tr>
</table>
</form>
<%
Else
Response.Write WriteMsg(no_entry)
End If
End If
End Sub
Sub repreg

IF DateDiff("N" ,Session("FloodTimer"),Now()) > strFloodTime Then
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
yazan = Session("Member")
amsg = Request.Form("amsg")
scode = duz(Request.Form("f2_sec_code"))
abone = Request.Form("abone")

If yazan = "" then
response.write WriteMsg("<p class=td>Sayfada cok uzun zaman beklemissiniz ve siteden cikmis gorunuyorsunuz, ya ust taraftan tekrar giris yapin yada alt kisimi kullanarak geri gidin ve yazdiklarinizi kopyalayin.</p>")
elseIF amsg="" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&f_error2&"</div><hr color="&tablo_cerceve&">"
ElseIF scode = "" Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&sc_err1&"<b></b></div><hr color="&tablo_cerceve&">"
ElseIF scode <> Session("Forum2SecurityCode") Then
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu>"&sc_err2&"<b></b></div><hr color="&tablo_cerceve&">"
%>
<!--#include file="inc/geri.asp" -->
<%
Else
FOR I = 1 TO UBound(saryHTMLtags)
amsg = Replace(amsg,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1)
Next
Set ms = Connection.Execute("Select * FROM f_mesajlar WHERE mesajid = "&msgid&"")
Set rent = Server.CreateObject("ADODB.RecordSet")
rSQL = "Select * FROM f_mesajlar"
rent.open rSQL,Connection,3,3
rent.AddNew
rent("mmsg") = amsg
rent("yazan") = yazan
rent("tarih") = Now()
rent("okunma") = msgid
rent("kategoriid") = ms("kategoriid")
rent.Update
If abone = "ol" then
Set aboneol = Server.CreateObject("ADODB.RecordSet")
rSQL = "Select * FROM f_abone"
aboneol.open rSQL,Connection,3,3
aboneol.AddNew
aboneol("gk") = uyebilgi("gk")
aboneol("kategoriid") = ms("kategoriid")
aboneol.Update
End if
Session("FloodTimerStart") = "Yes"
Set Rmsg = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&Session("Member")&"'")
Session("FloodTimer") = Rmsg("son_tarih")
Set msgUCount = Connection.Execute("UPDATE MEMBERS SET msg_sayisi = "&Rmsg("msg_sayisi")&"+1 WHERE kul_adi='"&Session("Member")&"'")
Response.Redirect "?action=Topic&id="&msgid&""
END IF
ELSE
Response.Write "<hr color="&tablo_cerceve&"><div align=center class=td_menu><b>"&f_error6&"</b></div><hr color="&tablo_cerceve&">"
End IF
End Sub
Sub Control
kategoriid = Request.QueryString("id")
kategoriid = QS_CLEAR(kategoriid)
Set cat = Connection.Execute("Select * FROM f_kategoriler WHERE kategoriid = "&kategoriid&"")
If Session("Level") <> "1" AND Session("Level") <> "2" AND Session("Level") <> "3" AND Session("Level") <> "4" AND Session("Level") <> "5" Then
If cat("noentry") = True Then
Response.Write "<BR><BR><BR><BR><div class=td_menu align=center>Forumumuzun bu konusu ve içeriðindeki mesajlar üye ve ziyaretçilerimize yasaklanmýþtýr, bu alana sadece yetkililer girebilir...</div><BR>"
%>
<CENTER><!--#include file="inc/geri.asp" --></CENTER><BR><BR><BR>
<%
ELSE
Application("Per") = "OK"
Response.Redirect "?action=Topics&id="&kategoriid&""
END IF
ELSE
Application("Per") = "OK"
Response.Redirect "?action=Topics&id="&kategoriid&""
END IF
End Sub
%>
<br>
<table border="0" cellspacing="0" cellpadding="2" width="100%" align="center">
	<tr>
		<td width="78%" valign="top">
			<table border="0" cellspacing="1" cellpadding="1" width="100%" bgcolor="<%=content_bg%>" class="td_menu">
				<tr>
					<td width="60%" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="18"><b>En Son 10 Konu</b></td>
					<td width="10%" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="18" align="center"><b><%=fr_pop5%></b></td>
					<td width="30%" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="18" align="center"><b><%=fw_pop5%></b></td>
				</tr>
<% Set pop5 = Connection.Execute("Select * FROM f_mesajlar WHERE topic = True ORDER BY tarih DESC")
For i = 1 TO 10
If pop5.Eof Then exit FOR
Set popm = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&pop5("yazan")&"'")
Set reply_count = Connection.Execute("Select Count(*) AS count FROM f_mesajlar WHERE okunma = "&pop5("mesajid")&" AND topic = False")
%>
				<tr bgcolor="<%=content_bg%>">
					<td width="60%"  ><a href="?action=Topic&id=<%=pop5("mesajid")%>"><b><%=duz(pop5("baslik"))%></b></a><br></td>
					<td width="10%" align="center" ><%=reply_count("count")%></td>
					<td width="30%" align="center" ><a href="members.asp?action=member_details&uid=<%=popm("uye_id")%>"><%=pop5("yazan")%></a></td>
				</tr>
<%
pop5.MoveNext
Next
%>
			</table>
		</td>
		<td width="22%" valign="top">
			<table border="0" cellspacing="1" cellpadding="1" width="100%" class="td_menu">
				<tr>
					<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="18">
					<b><CENTER><%=f_stats%></CENTER></b></td>
				</tr>
				<tr bgcolor="<%=content_bg%>">
					<td>
					<%
					Set mforum = Connection.Execute("Select Count(*) AS count FROM f_anakategori")
					Set forum = Connection.Execute("Select Count(*) AS count FROM f_kategoriler")
					Set msgs = Connection.Execute("Select Count(*) AS count FROM f_mesajlar WHERE topic = True")
					Set reps = Connection.Execute("Select Count(*) AS count FROM f_mesajlar WHERE topic = False")
					%>
					<b><%=sett("site_isim")%></b> Forumlarýnda <B><%=mforum("count")%></B> kategoride , <b><%=forum("count")%></b> forum ve bu forumlarýn içerisinde <b><%=msgs("count")%></b> tane konu , <b><%=reps("count")%></b> tane yanýt bulunmaktadýr.
					</td>
				</tr>
			</table>
<BR>
			<table border="0" cellspacing="1" cellpadding="1" width="100%" class="td_menu" bgcolor="<%=content_bg%>">
				<tr>
					<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="18">
					<b><CENTER>Forumda Ara</CENTER></b></td>
				</tr>
				<tr>
					<td>
<%Set fRs = Connection.Execute("Select * FROM f_kategoriler ORDER BY katad ASC")%>
						<table border="0" cellspacing="0" cellpadding="0" class="td_menu" width="100%">
<form method="post" action="?action=Results">
							<tr bgcolor="<%=content_bg%>">
								<td width="50%" align="right"><%=fs_word%> : </td>
								<td width="50%"><input type="text" name="word" size="30" class="form1"></td>
							</tr>
							<tr bgcolor="<%=content_bg%>">
								<td width="50%" align="right"><%=fs_forums%> : </td>
								<td width="50%">
								<select name="forums" size="1" class="form1">
								<% Do Until fRs.Eof %>
								<option value="<%=fRs("kategoriid")%>"><%=fRs("katad")%></option>
								<%
								fRs.MoveNext
								Loop
								%>
								</select>
								</td>
							</tr>
							<tr bgcolor="<%=content_bg%>">
								<td width="50%" align="right"></td>
								<td width="50%"><input type="submit" value="   <%=fs_button%>   " class="submit"></td>
							</tr>
						</form>
						</table>
<%
fRs.Close
Set fRs = Nothing
%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu" bgcolor="<%=content_bg%>">
<tr>
<td background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="20"><%=f_whoonline%></td>
</tr>
<tr bgcolor="<%=content_bg%>">
<td>
<%
Session.LCID = 1033
liste = DateAdd("n", -1 * 1, Now())
Set guests = Connection.Execute("SELECT COUNT(*) AS sayi FROM GUESTS")
Set om = Connection.Execute("SELECT * FROM MEMBERS where son_tarih >= #"&liste&"# ORDER BY son_tarih DESC")
Set w_online = Connection.Execute("UPDATE MEMBERS SET u_online = True where son_tarih >= #"&liste&"#")
Set w_offline = Connection.Execute("UPDATE MEMBERS SET u_online = False where son_tarih <= #"&liste&"#")

Set mem_c = Connection.Execute("Select COUNT(*) AS say FROM MEMBERS where son_tarih >= #"&liste&"#")

IF om.eof Then
mem_count = "0"
Else
mem_count = mem_c("say")
End If
%>
<%
IF om.EoF Then
Response.Write no_online
ELSE
Do Until om.Eof
IF om("seviye") = "1" Then
om_title = ""&level1&""
ElseIf om("seviye") = "2" Then
om_title = ""&level2&""
ElseIf om("seviye") = "3" Then
om_title = ""&level3&""
ElseIf om("seviye") = "4" Then
om_title = ""&level4&""
ElseIf om("seviye") = "5" Then
om_title = ""&level5&""
ElseIf om("seviye") = "0" Then
om_title = ""&normal&""
End If
IF om("seviye") = "1" Then
om_t_c = online_admin
om_t_style = "bold"
ElseIF om("seviye") = "2" OR om("seviye") = "3" OR om("seviye") = "4" Then
om_t_c = online_editor
om_t_style = "bold"
ELSE
om_t_style = "regular"
End IF
Response.Write "<a href=""members.asp?action=member_details&uid="&om("uye_id")&""" title="""&om_title&"""><font style=""font-weight:"&om_t_style&""">"&om("kul_adi")&"</a> , "
om.MoveNext
Loop
End IF
%>
</td>
</tr>
</table>
<% call ALT %>                   