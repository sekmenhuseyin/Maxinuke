<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#											Resmim Blok Uygulamasi V 1.0											#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
'IF Session("Enter") = "Yes" Then
'	If uyebilgi("resmim_onay")=true Then
'	resim=uyebilgi("resmim")
'	Else
'	resim="yok.gif"
'	End if
'else
'	resim="yok.gif"
'End if
	%>
	<a ONCLICK="window.open('upload_r.asp','upload_r','top=20,left=20,width=350,height=100,toolbar=no,scrollbars=yes');"><IMG SRC="uploads/avatar/<%'=resim%>" alt="Degistimek icin tiklayin" WIDTH="140" BORDER="0" align="left"></A>