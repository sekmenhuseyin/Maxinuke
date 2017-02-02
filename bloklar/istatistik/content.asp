<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Istatistikler Blok Uygulamasi V 2.1											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
IF Request.Cookies("MaxiNuke_Counter")("C_ENTER") <> "Yes" AND Request.Cookies("MaxiNuke_Counter")("C_DATE") <> Date() Then
Set count = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = Date()-0")
Set countAdd = Connection.Execute("UPDATE WEBCOUNTER SET today = "&count("today")&" + 1 WHERE cdate = Date()-0")
Response.Cookies("MaxiNuke_Counter")("C_ENTER") = "Yes"
Response.Cookies("MaxiNuke_Counter")("C_DATE") = Date()
Set countAdd = Nothing
count.Close
Set count = Nothing
End IF
liste = DateAdd("n", -5, Now())
Set mem = Connection.Execute("Select * FROM MEMBERS ORDER BY uye_id DESC")
Set memc = Connection.Execute("Select Count(*) AS count FROM MEMBERS")
Set counter = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = date()-0")
Set counter2 = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = date()-1")
Set counter3 = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = date()-2")
Set counter_t = Connection.Execute("SELECT SUM(today) AS count FROM WEBCOUNTER")
Set today = Connection.Execute("Select Count(*) AS m_count FROM MEMBERS WHERE uyelik_tarihi=date()-0")
Set yesterday = Connection.Execute("Select Count(*) AS m_count FROM MEMBERS WHERE uyelik_tarihi=date()-1")
Set evvelsigun = Connection.Execute("Select Count(*) AS m_count FROM MEMBERS WHERE uyelik_tarihi=date()-2")
Set guests = Connection.Execute("SELECT COUNT(*) AS sayi FROM GUESTS")
Set om = Connection.Execute("SELECT * FROM MEMBERS where son_tarih >= #"&liste&"# ORDER BY son_tarih DESC")
Set w_online = Connection.Execute("UPDATE MEMBERS SET u_online = True where son_tarih >= #"&liste&"#")
Set w_offline = Connection.Execute("UPDATE MEMBERS SET u_online = False where son_tarih <= #"&liste&"#")
Set mem_c = Connection.Execute("Select COUNT(*) AS say FROM MEMBERS where son_tarih >= #"&liste&"#")
If sett("onaylama") = True then
Set oks = Connection.Execute("Select Count(*) AS m_count FROM MEMBERS WHERE onay=false")
End if
IF counter2.EoF OR counter2.BoF Then
counter2_yesterday = "0"
ELSE
counter2_yesterday = counter2("today")
End IF
IF counter3.EoF OR counter3.BoF Then
counter3_yesterday = "0"
ELSE
counter3_yesterday = counter3("today")
End IF
IF om.eof Then
mem_count = "0"
Else
mem_count = mem_c("say")
End If
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" class="td_menu">
<tr>
<td><img src="IMAGES/site/mem.gif"> <b><%=menu4%></b></td>
</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=last_member%>:<a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><b><%=mem("kul_adi")%></b></a></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=c_today%>:<b><%=today("m_count")%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=c_yesterday%>:<b><%=yesterday("m_count")%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> Evvelsi Gün:<b><%=evvelsigun("m_count")%></b></td>
	</tr>
<%If sett("onaylama") = True then%>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> Onaysiz:<b><%=oks("m_count")%></b></td>
	</tr>
<%End if%>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=total_member%>:<b>1<%=memc("count")%></b></td>
	</tr>
	<tr>
		<td align="center"><hr size="1" color="<%=tablo_cerceve%>"></td></tr>
	<tr>
	<tr>
		<td><img src="IMAGES/site/visitors.gif"> <b><a href="Members.asp?action=OnlineUsers"><%=sitedekiler%></a></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=bottom_members%>:<b><%=mem_count%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=bottom_ziyaretciler%>:<b><%=guests("sayi")%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=bottom_total%>:<b><%=mem_count+guests("sayi")%></b></td>
	</tr>
	<tr>
		<td align="center"><hr size="1" color="<%=tablo_cerceve%>"></td></tr>
	<tr>
	<tr>
		<td><img src="IMAGES/site/online.gif"> <b><%=online_mems%></b></td>
	</tr>
	<tr>
		<td>
<%
num = "0"
Do Until om.Eof
num = num+1
Response.Write ""&num&": <a href=""members.asp?action=member_details&uid="&om("uye_id")&""">"&om("kul_adi")&"</a><br>"
om.MoveNext
Loop
IF num = "0" Then
Response.Write no_online
End IF
%>
		</td>
	</tr>
	<tr>
		<td align="center"><hr size="1" color="<%=tablo_cerceve%>"></td>
	</tr>
	<tr>
		<td><img src="IMAGES/site/graph.gif"> <b><%=site_hits%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=c_today%>&nbsp;:&nbsp;<b><%=counter("today")%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=c_yesterday%>&nbsp;:&nbsp;<b><%=counter2_yesterday%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> Evvelsi Gün&nbsp;:&nbsp;<b><%=counter3_yesterday%></b></td>
	</tr>
	<tr>
		<td><img src="images/temalar/<%=Session("SiteTheme")%>/arrow.gif"> <%=c_total%>&nbsp;:&nbsp;<b><%=counter_t("count")%></b></td>
	</tr>
</table>
<%
mem_c.Close
Set mem_c = Nothing
Set w_offline = Nothing
Set w_online = Nothing
om.Close
Set om = Nothing
guests.Close
Set guests = Nothing
If sett("onaylama") = True Then
oks.Close
Set oks = Nothing
End if
evvelsigun.Close
Set evvelsigun = Nothing
yesterday.Close
Set yesterday = Nothing
today.Close
Set today = Nothing
counter_t.Close
Set counter_t = Nothing
counter3.Close
Set counter3 = Nothing
counter2.Close
Set counter2 = Nothing
counter.Close
Set counter = Nothing
memc.Close
Set memc = Nothing
mem.Close
Set mem = Nothing
%>