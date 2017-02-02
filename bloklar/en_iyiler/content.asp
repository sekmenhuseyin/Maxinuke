<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											En Iyiler Blok Uygulamasi V 2.1											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
Set aktif_uye = Connection.Execute("Select * FROM MEMBERS ORDER BY login_sayisi DESC")
Set aktif_yazi = Connection.Execute("Select * FROM makale WHERE a_approved = True ORDER BY hit DESC")
Set aktif_dosya = Connection.Execute("Select * FROM DOWNLOADS WHERE onay = True ORDER BY p_hit DESC")
Set aktif_forum = Connection.Execute("Select * FROM f_mesajlar WHERE kilit = false and topic=true ORDER BY okunma DESC")
%>
<TABLE class="td_menu" width="100%">
<TR>
	<TD colspan="2" align="center" ><b>En Hýzlý 5 Üyemiz</b></TD>
</TR>
<%
For I = 1 TO 5
if aktif_uye.eof Then exit For
%>
<TR>
	<TD><%=I%></TD>
	<TD><a href="members.asp?action=member_details&uid=<%=aktif_uye("uye_id")%>"><%=aktif_uye("kul_adi")%></a></TD>
</TR>
<%
aktif_uye.MoveNext
Next
aktif_uye.Close
Set aktif_uye = Nothing
%>
<tr>
	<TD colspan="2"><hr size="1" color="<%=tablo_cerceve%>"></TD>
<TR>
<TR>
	<TD colspan="2" align="center" ><b>En Çok Okunan 5 Yazý</b></TD>
<%
For p = 1 To 5
if aktif_yazi.eof Then exit For
%>
</TR>
	<TD><%=p%></TD>
	<TD><a href="makale.asp?action=p_details&aid=<%=aktif_yazi("aid")%>"><%=aktif_yazi("a_title")%></a></TD>
</TR>
<%
aktif_yazi.MoveNext
Next
aktif_yazi.Close
Set aktif_yazi = Nothing
%>
<tr>
	<TD colspan="2"><hr size="1" color="<%=tablo_cerceve%>"></TD>
<TR>
<TR>
	<TD colspan="2" align="center" ><b>En Populer 5 Dosya</b></TD>
<%
For m = 1 To 5
if aktif_dosya.eof Then exit For
%>
</TR>
	<TD><%=m%></TD>
	<TD><a href="Programs.asp?action=p_details&pid=<%=aktif_dosya("pid")%>"><%=aktif_dosya("p_name")%></a></TD>
</TR>
<%
aktif_dosya.MoveNext
Next
aktif_dosya.Close
Set aktif_dosya = Nothing
%>
<tr>
	<TD colspan="2"><hr size="1" color="<%=tablo_cerceve%>"></TD>
<TR>
<TR>
	<TD colspan="2" align="center" ><b>En Populer 5 Forum</b></TD>
<%
For h = 1 To 5
if aktif_forum.eof Then exit For
%>
</TR>
	<TD><%=h%></TD>
	<TD><a href="forum.asp?action=Topic&id=<%=aktif_forum("mesajid")%>"><%=aktif_forum("baslik")%></a></TD>
</TR>
<%
aktif_forum.MoveNext
Next
aktif_forum.Close
Set aktif_forum = Nothing
%>
</TABLE>