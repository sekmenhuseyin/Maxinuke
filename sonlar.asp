<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												  Sonlar Uygulamasi													#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
--><BR>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
<TR>
	<TD>
		<%
		If sett("sd") = True Then
		set dosya_5 = server.createobject("adodb.recordset")
		SQL = "Select * from DOWNLOADS where onay=true order by pid desc"
		dosya_5.open SQL,Connection,1,3
		%>
		<table width="100%" height="100" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
			<TR>
				<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="Programs.asp?action=cats">SON DOSYALAR</A></B></TD>
			</TR>
			<TR>
				<TD bgcolor="<%=content_bg%>" valign="top" style="padding-left:3px;padding-right:3px">
				<%
				if dosya_5.eof then
				response.Write "<BR><BR><CENTER><B>Henüz Dosya Eklenmemiþ</B></CENTER>"
				else
				for k=1 to 5
				if dosya_5.eof or dosya_5.bof then exit for
				%>
				<IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
				<A HREF="Programs.asp?action=p_details&pid=<%=dosya_5("pid")%>"><%=Left(dosya_5("p_name"),45)%></A><br>
				<%
				dosya_5.movenext : Next
				End if
				%>
				</TD>
			</TR>
		</TABLE>
		<%
		dosya_5.close
		set dosya_5=Nothing
		End if
		%>
	</TD>
	<TD>
		<%
		If sett("sm") = True Then
		set makale_5 = server.createobject("adodb.recordset")
		SQL = "Select * from makale where a_approved=true order by aid desc"
		makale_5.open SQL,Connection,1,3
		%>
		<table width="100%" height="100" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
			<TR>
				<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="makale.asp?action=cats">SON MAKALELER</A></B></TD>
			</TR>
			<TR>
				<TD bgcolor="<%=content_bg%>" valign="top" style="padding-left:3px;padding-right:3px">
				<%
				if makale_5.eof then
				response.Write "<BR><BR><CENTER><B>Henüz Makale Eklenmemiþ</B></CENTER>"
				else
				for k=1 to 5
				if makale_5.eof or makale_5.bof then exit for
				%>
				<IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
				<A HREF="makale.asp?action=p_details&aid=<%=makale_5("aid")%>"><%=Left(makale_5("a_title"),45)%></A><%If Len(makale_5("a_title"))>45 Then%>...<%End If%><br>
				<%
				makale_5.movenext : Next
				End if
				%>
				</TD>
			</TR>
		</TABLE>	
		<%
		makale_5.close
		set makale_5=Nothing
		End if
		%>
	</TD>
	<TD>
		<%
		If sett("su") = True Then
		set uye_5 = server.createobject("adodb.recordset")
		SQL = "Select * from MEMBERS ORDER BY uye_id DESC"
		uye_5.open SQL,Connection,1,3
		%>
		<table width="100%" height="100" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
			<TR>
				<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="Members.asp?action=members">SON ÜYELER</A></B></TD>
			</TR>
			<TR>
				<TD bgcolor="<%=content_bg%>" valign="top" style="padding-left:3px;padding-right:3px">
				<%
				for k=1 to 5
				if uye_5.eof or uye_5.bof then exit for
				%>
				<IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
				<A HREF="members.asp?action=member_details&uid=<%=uye_5("uye_id")%>"><%=Left(uye_5("kul_adi"),45)%></A><br>
				<%uye_5.movenext : next%>
				</TD>
			</TR>
		</TABLE>	
		<%
		uye_5.close
		set uye_5=Nothing
		End if
		%>
	</TD>
</TR>
</table>
<BR>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
<TR>
	<TD>
		<%
		If sett("sf") = True Then
		set fikra_5 = server.createobject("adodb.recordset")
		SQL = "Select * from fikra where onay=true ORDER BY id DESC"
		fikra_5.open SQL,Connection,1,3
		%>
		<table width="100%" height="100" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
			<TR>
				<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="moduller.asp?mi=13">SON FIKRALAR</A></B></TD>
			</TR>
			<TR>
				<TD bgcolor="<%=content_bg%>" valign="top" style="padding-left:3px;padding-right:3px">
				<%
				if fikra_5.eof then
				response.Write "<BR><BR><CENTER><B>Henüz Fýkra Eklenmemiþ</B></CENTER>"
				else
				for k=1 to 5
				if fikra_5.eof or fikra_5.bof then exit for
				%>
				<IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
				<A HREF="moduller.asp?mi=13&action=oku&id=<%=fikra_5("id")%>"><%=Left(fikra_5("baslik"),45)%></A><br>
				<%fikra_5.movenext : Next
				End if
				%>
				</TD>
			</TR>
		</TABLE>	
		<%
		fikra_5.close
		set fikra_5=Nothing
		End if
		%>
	</TD>
	<TD>
		<%
		If sett("sfb") = True Then
		set forum_5 = server.createobject("adodb.recordset")
		SQL = "Select * from f_mesajlar WHERE topic = True ORDER BY mesajid DESC"
		forum_5.open SQL,Connection,1,3
		%>
		<table width="100%" height="100" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
			<TR>
				<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="Forum.asp">SON FORUM BASLIKLARI</A></B></TD>
			</TR>
			<TR>
				<TD bgcolor="<%=content_bg%>" valign="top" style="padding-left:3px;padding-right:3px">
				<%
				if forum_5.eof then
				response.Write "<BR><BR><CENTER><B>Henüz Forum Eklenmemiþ</B></CENTER>"
				else
				for k=1 to 5
				if forum_5.eof or forum_5.bof then exit for
				%>
				<IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
				<A HREF="forum.asp?action=Topic&id=<%=forum_5("mesajid")%>"><%=Left(forum_5("baslik"),45)%></A><br>
				<%forum_5.movenext : Next
				End if
				%>
				</TD>
			</TR>
		</TABLE>	
		<%
		forum_5.close
		set forum_5=Nothing
		End if
		%>
	</TD>
	<TD>
<%
If sett("sr") = True Then
set resim_5 = server.createobject("adodb.recordset")
SQL = "Select * from rg ORDER BY id DESC"
resim_5.open SQL,Connection,1,3
%>
		<table width="100%" height="100" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="moduller.asp?mi=7">SON RESIMLER</A></B></TD>
				</TR>
				<TR>
				<TD bgcolor="<%=content_bg%>" valign="top" style="padding-left:3px;padding-right:3px">
					<%
					if resim_5.eof then
					response.Write "<BR><BR><CENTER><B>Henüz Resim Eklenmemiþ</B></CENTER>"
					else
					for k=1 to 5
					if resim_5.eof or resim_5.bof then exit for
					%>
					<IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif">
					<A HREF="moduller.asp?mi=7&action=bak&id=<%=resim_5("id")%>"><%=Left(resim_5("baslik"),45)%></A><br>
					<%resim_5.movenext : Next
					End if
					%>
					</TD>
				</TR>
			</TABLE>
<%
resim_5.close
set resim_5=Nothing
End If
%>
	</TD>
</TR>
</TABLE>
<BR>
<%
If sett("so") = True Then
set oyun_5 = server.createobject("adodb.recordset")
SQL = "Select * from fo ORDER BY id desc"
oyun_5.open SQL,Connection,1,3
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<TR>
		<TD valign="top" width="33%">
			<table width="100%" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="moduller.asp?mi=10">SON OYUNLAR</A></B></TD>
				</TR>
				<TR>
					<TD bgcolor="<%=content_bg%>">
					<%
					if oyun_5.eof then
					response.Write "<BR><CENTER><B>Henüz Oyun Eklenmemiþ</B></CENTER>"
					End If
					%>
					<marquee scrollamount="3" scrolldelay="70" onmouseover="this.stop();" onmouseout="this.start();" direction="right">
					<%
					for k=1 to 20
					if oyun_5.eof or oyun_5.bof then exit for
					%>
					<A HREF="moduller.asp?mi=10&action=bak&id=<%=oyun_5("id")%>"><IMG SRC="uploads/image/fo/<%=oyun_5("resim")%>" WIDTH="115" HEIGHT="90" BORDER="0" ALT="<%=oyun_5("baslik")%>"></A>&nbsp;&nbsp;<%oyun_5.movenext : Next%></marquee>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<%
oyun_5.close
set oyun_5=Nothing
End if
%>
<BR>
<%
If sett("se") = True Then
set ekart_5 = server.createobject("adodb.recordset")
SQL = "Select * from ekartlar where onay=true and tip=1 ORDER BY id DESC"
ekart_5.open SQL,Connection,1,3
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<TR>
		<TD valign="top" width="33%">
			<table width="100%" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B><A HREF="moduller.asp?mi=12">SON E-KARTLAR</A></B></TD>
				</TR>
				<TR>
					<TD bgcolor="<%=content_bg%>">
					<%
					if ekart_5.eof then
					response.Write "<CENTER><BR><BR><BR><B>Kayitli E-Kart Yok</B><BR><BR><BR></CENTER>"
					End if
					%>
					<marquee scrollamount="3" scrolldelay="100" onmouseover="this.stop();" onmouseout="this.start();">
					<%
					for k=1 to 20

					if ekart_5.eof or ekart_5.bof then exit for

					%>
					<A HREF="moduller.asp?mi=12&eylem=olustur&id=<%=ekart_5("id")%>"><IMG SRC="uploads/image/ekart/<%=ekart_5("yol")%>" WIDTH="115" HEIGHT="90" BORDER="0"></A>&nbsp;&nbsp;<%ekart_5.movenext : next%></marquee>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<%
ekart_5.close
set ekart_5=Nothing
End if
%>
<BR>
<%
If sett("sv") = True Then
set video_5 = server.createobject("adodb.recordset")
SQL = "Select * from videolar where onay=true ORDER BY id DESC"
video_5.open SQL,Connection,1,3
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<TR>
		<TD valign="top" width="33%">
			<table width="100%" border="0" cellspacing="3" cellpadding="0" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B>SON VIDEOLAR</B></TD>
				</TR>
				<TR>
					<TD bgcolor="<%=content_bg%>">
					<%
					if video_5.eof then
					response.Write "<CENTER><BR><BR><BR><B>Kayitli video Yok</B><BR><BR><BR></CENTER>"
					End if
					%>
					<marquee scrollamount="3" scrolldelay="100" onmouseover="this.stop();" onmouseout="this.start();">
					<%
					for k=1 to 20

					if video_5.eof or video_5.bof then exit for

					%>
					<a href="moduller.asp?mi=16&eylem=olustur&id=<%=video_5("id")%>"><IMG SRC="http://i.ytimg.com/vi/<%=video_5("yol")%>/default.jpg"  width="120" height="90" BORDER="0"></a>&nbsp;&nbsp;<%video_5.movenext : next%></marquee>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<%
video_5.close
set video_5=Nothing
End if
%>