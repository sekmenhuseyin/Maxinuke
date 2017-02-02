<!--#include file="../../INC/includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#										Site Kadrosu Modul Uygulamasi V 2.0											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
set kadro1 = server.createobject("adodb.recordset")
SQL = "Select * from members where seviye=1 order by uye_id"
kadro1.open SQL,Connection,1,3
set kadro2 = server.createobject("adodb.recordset")
SQL = "Select * from members where seviye=2 order by uye_id"
kadro2.open SQL,Connection,1,3
set kadro3 = server.createobject("adodb.recordset")
SQL = "Select * from members where seviye=3 order by uye_id"
kadro3.open SQL,Connection,1,3
set kadro4 = server.createobject("adodb.recordset")
SQL = "Select * from members where seviye=4 order by uye_id"
kadro4.open SQL,Connection,1,3
set kadro5 = server.createobject("adodb.recordset")
SQL = "Select * from members where seviye=5 order by uye_id"
kadro5.open SQL,Connection,1,3
%>
<table width="100%" border="0" cellspacing="5" cellpadding="5" align="center">
	<TR>
		<TD>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B>GENEL YÖNETÝCÝLERÝMÝZ</B></TD>
				</TR>
			</table>
<%
for k=1 to kadro1.RecordCount
if kadro1.eof or kadro1.bof then exit For
	If kadro1("resmim_onay")=true Then
	resim=kadro1("resmim")
	Else
	resim="yok.gif"
	End if
%>
			<table width="100%" border="0" cellspacing="3" cellpadding="3" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD><A HREF="members.asp?action=member_details&uid=<%=kadro1("uye_id")%>"><IMG SRC="uploads/avatar/<%=resim%>" WIDTH="120" BORDER="0" align="left"><CENTER><B>-=-=-=-=-=-=-=-=-=-=-=-=- <%=kadro1("kul_adi")%> -=-=-=-=-=-=-=-=-=-=-=-=-</B></CENTER>Site yöneticilerimizden <%=kadro1("kul_adi")%>;&nbsp;<%=kadro1("yas")%> doðumlu, <%=kadro1("sehir")%> bölgesinde ikamet etmekte ve <%=kadro1("meslek")%> ayný zamanda <%=kadro1("uyelik_tarihi")%> tarihinden bu yana sitemizde.</A>
					</TD>
				</TR>
			</TABLE>		
<%kadro1.movenext : next%>  
		</TD>
	</TR>
</TABLE>
<!--#include file="../../inc/adsense_3.asp" -->
<table width="100%" border="0" cellspacing="5" cellpadding="5" align="center">
	<TR>
		<TD>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B>DOWNLOAD / DOSYA / PROGRAM EDÝTÖRLERÝMÝZ</B></TD>
				</TR>
			</table>
<%
for k=1 to kadro2.RecordCount
if kadro2.eof or kadro2.bof then exit For
	If kadro2("resmim_onay")=true Then
	resim=kadro2("resmim")
	Else
	resim="yok.gif"
	End if
%>
			<table width="100%" border="0" cellspacing="3" cellpadding="3" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD><A HREF="members.asp?action=member_details&uid=<%=kadro2("uye_id")%>"><IMG SRC="uploads/avatar/<%=resim%>" WIDTH="120" BORDER="0" align="left"><CENTER><B>-=-=-=-=-=-=-=-=-=-=-=-=- <%=kadro2("kul_adi")%> -=-=-=-=-=-=-=-=-=-=-=-=-</B></CENTER>Site yöneticilerimizden <%=kadro2("kul_adi")%>;&nbsp;<%=kadro2("yas")%> doðumlu, <%=kadro2("sehir")%> bölgesinde ikamet etmekte ve <%=kadro2("meslek")%> ayný zamanda <%=kadro2("uyelik_tarihi")%> tarihinden bu yana sitemizde.</A>
					</TD>
				</TR>
			</TABLE>		
<%kadro2.movenext : next%>  
		</TD>
	</TR>
</TABLE>
<!--#include file="../../inc/adsense_3.asp" -->
<table width="100%" border="0" cellspacing="5" cellpadding="5" align="center">
	<TR>
		<TD>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B>MAKALE / YAZI EDÝTÖRLERÝMÝZ</B></TD>
				</TR>
			</table>
<%
for k=1 to kadro3.RecordCount
if kadro3.eof or kadro3.bof then exit For
	If kadro3("resmim_onay")=true Then
	resim=kadro3("resmim")
	Else
	resim="yok.gif"
	End if
%>
			<table width="100%" border="0" cellspacing="3" cellpadding="3" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD><A HREF="members.asp?action=member_details&uid=<%=kadro3("uye_id")%>"><IMG SRC="uploads/avatar/<%=resim%>" WIDTH="120" BORDER="0" align="left"><CENTER><B>-=-=-=-=-=-=-=-=-=-=-=-=- <%=kadro3("kul_adi")%> -=-=-=-=-=-=-=-=-=-=-=-=-</B></CENTER>Site yöneticilerimizden <%=kadro3("kul_adi")%>;&nbsp;<%=kadro3("yas")%> doðumlu, <%=kadro3("sehir")%> bölgesinde ikamet etmekte ve <%=kadro3("meslek")%> ayný zamanda <%=kadro3("uyelik_tarihi")%> tarihinden bu yana sitemizde.</A>
					</TD>
				</TR>
			</TABLE>		
<%kadro3.movenext : next%>  
		</TD>
	</TR>
</TABLE>
<!--#include file="../../inc/adsense_3.asp" -->
<table width="100%" border="0" cellspacing="5" cellpadding="5" align="center">
	<TR>
		<TD>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B>HABER EDÝTÖRLERÝMÝZ</B></TD>
				</TR>
			</table>
<%
for k=1 to kadro4.RecordCount
if kadro4.eof or kadro4.bof then exit For
	If kadro4("resmim_onay")=true Then
	resim=kadro4("resmim")
	Else
	resim="yok.gif"
	End if
%>
			<table width="100%" border="0" cellspacing="3" cellpadding="3" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD><A HREF="members.asp?action=member_details&uid=<%=kadro4("uye_id")%>"><IMG SRC="uploads/avatar/<%=resim%>" WIDTH="120" BORDER="0" align="left"><CENTER><B>-=-=-=-=-=-=-=-=-=-=-=-=- <%=kadro4("kul_adi")%> -=-=-=-=-=-=-=-=-=-=-=-=-</B></CENTER>Site yöneticilerimizden <%=kadro4("kul_adi")%>;&nbsp;<%=kadro4("yas")%> doðumlu, <%=kadro4("sehir")%> bölgesinde ikamet etmekte ve <%=kadro4("meslek")%> ayný zamanda <%=kadro4("uyelik_tarihi")%> tarihinden bu yana sitemizde.</A>
					</TD>
				</TR>
			</TABLE>		
<%kadro4.movenext : next%>  
		</TD>
	</TR>
</TABLE>
<!--#include file="../../inc/adsense_2.asp" -->
<table width="100%" border="0" cellspacing="5" cellpadding="5" align="center">
	<TR>
		<TD>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
				<TR>
					<TD align="center" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg_lite.gif" height="20"><B>FORUM OPERATÖR VE  EDÝTÖRLERÝMÝZ</B></TD>
				</TR>
			</table>
<%
for k=1 to kadro5.RecordCount
if kadro5.eof or kadro5.bof then exit For
	If kadro5("resmim_onay")=true Then
	resim=kadro5("resmim")
	Else
	resim="yok.gif"
	End if
%>
			<table width="100%" border="0" cellspacing="3" cellpadding="3" align="center" class="td_menu" style="<%=TableBorder%>">
				<TR>
					<TD><A HREF="members.asp?action=member_details&uid=<%=kadro5("uye_id")%>"><IMG SRC="uploads/avatar/<%=resim%>" WIDTH="120" BORDER="0" align="left"><CENTER><B>-=-=-=-=-=-=-=-=-=-=-=-=- <%=kadro5("kul_adi")%> -=-=-=-=-=-=-=-=-=-=-=-=-</B></CENTER>Site yöneticilerimizden <%=kadro5("kul_adi")%>;&nbsp;<%=kadro5("yas")%> doðumlu, <%=kadro5("sehir")%> bölgesinde ikamet etmekte ve <%=kadro5("meslek")%> ayný zamanda <%=kadro5("uyelik_tarihi")%> tarihinden bu yana sitemizde.</A>
					</TD>
				</TR>
			</TABLE>		
<%kadro5.movenext : next%>  
		</TD>
	</TR>
</TABLE>
<!--#include file="../../inc/adsense_2.asp" -->
<%
kadro5.close
set kadro5=nothing
kadro4.close
set kadro4=nothing
kadro3.close
set kadro3=nothing
kadro2.close
set kadro2=nothing
kadro1.close
set kadro1=nothing
Connection.close
set Connection=nothing
%>