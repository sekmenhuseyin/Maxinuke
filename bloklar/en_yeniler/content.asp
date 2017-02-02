<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											En Yeniler Blok Uygulamasi V 2.1										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
if sett("onaylama") = True then
Set yeni_uye = Connection.Execute("Select * FROM MEMBERS where onay=true ORDER BY uye_id DESC")
else
Set yeni_uye = Connection.Execute("Select * FROM MEMBERS ORDER BY uye_id DESC")
End if
Set yeni_yazi = Connection.Execute("Select * FROM makale WHERE a_approved = True ORDER BY aid DESC")
Set yeni_dosya = Connection.Execute("Select * FROM DOWNLOADS WHERE onay = True ORDER BY pid DESC")
Set yeni_forum = Connection.Execute("Select * FROM f_mesajlar WHERE kilit = false and topic=true ORDER BY mesajid DESC")
%>
<TABLE class="td_menu" width="100%">
<TR>
	<TD colspan="2" align="center" ><b>En Yeni 5 Üyemiz</b></TD>
</TR>
<%
For q = 1 TO 5
if yeni_uye.eof Then exit For
%>
<TR>
	<TD><font face=webdings color="<%=text%>" size=1>4</font></TD>
	<TD><a href="members.asp?action=member_details&uid=<%=yeni_uye("uye_id")%>"><%=yeni_uye("kul_adi")%></a></TD>
</TR>
<%
yeni_uye.MoveNext
Next
yeni_uye.Close
Set yeni_uye = Nothing
%>
<tr>
	<TD colspan="2"><hr size="1" color="<%=td_border_color%>"></TD>
<TR>
<TR>
	<TD colspan="2" align="center" ><b>En Yeni 5 Yazý</b></TD>
<%
For w = 1 To 5
if yeni_yazi.eof Then exit For
%>
</TR>
	<TD><font face=webdings color="<%=text%>" size=1>4</font></TD>
	<TD><a href="makale.asp?action=Read&aid=<%=yeni_yazi("aid")%>"><%=yeni_yazi("a_title")%></a></TD>
</TR>
<%
yeni_yazi.MoveNext
Next
yeni_yazi.Close
Set yeni_yazi = Nothing
%>
<tr>
	<TD colspan="2"><hr size="1" color="<%=td_border_color%>"></TD>
<TR>
<TR>
	<TD colspan="2" align="center" ><b>En Yeni 5 Dosya</b></TD>
<%
For e = 1 To 5
if yeni_dosya.eof Then exit For
%>
</TR>
	<TD><font face=webdings color="<%=text%>" size=1>4</font></TD>
	<TD><a href="Programs.asp?action=p_details&pid=<%=yeni_dosya("pid")%>"><%=yeni_dosya("p_name")%></a></TD>
</TR>
<%
yeni_dosya.MoveNext
Next
yeni_dosya.Close
Set yeni_dosya = Nothing
%>
<tr>
	<TD colspan="2"><hr size="1" color="<%=td_border_color%>"></TD>
<TR>
<TR>
	<TD colspan="2" align="center" ><b>En Yeni 5 Forum</b></TD>
<%
For e = 1 To 5
if yeni_forum.eof Then exit For
%>
</TR>
	<TD><font face=webdings color="<%=text%>" size=1>4</font></TD>
	<TD><a href="forum.asp?action=Topic&id=<%=yeni_forum("mesajid")%>"><%=yeni_forum("baslik")%></a></TD>
</TR>
<%
yeni_forum.MoveNext
Next
yeni_forum.Close
Set yeni_forum = Nothing
%>
</TABLE>