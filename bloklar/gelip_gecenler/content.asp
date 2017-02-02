<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Gelip Gecenler Blok Uygulamasi V 2.0										#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%set gelipgecen = connection.execute("select * from members order by son_tarih desc")%>
<marquee onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="200" SCROLLAMOUNT="1" border="0">
<TABLE class="td_menu" cellspacing="0" cellpadding="0" border=0 width=100%>
<%
for sttr = 1 to 50
if gelipgecen.eof Then exit For
son = gelipgecen("son_tarih")
fark = datediff("n",son,now) & " dk."
if fark="0 dk." then fark="Online"
%>
	<TR>
		<TD><IMG SRC="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"></td>
		<TD>&nbsp;<a href="members.asp?action=member_details&uid=<%=gelipgecen("uye_id")%>"><%=gelipgecen("kul_adi")%></a> (<%=fark%>)</td>
	</TR>
<%
gelipgecen.movenext
next
%>
</TABLE>
</marquee>
<%
gelipgecen.Close
Set gelipgecen = Nothing
%>