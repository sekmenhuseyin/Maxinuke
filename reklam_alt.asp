<%
Set altreklam = Server.CreateObject("ADODB.Recordset" )
SQLkk = "Select * FROM BANNERS WHERE active=True AND position='bottom'"
altreklam.Open SQLkk,Connection,3,3
if altreklam.eof then
%>
<!--#include file="inc/adsense_3.asp" -->
<%
else
TotalRecord2 = altreklam.RecordCount
Randomize
MoveEntry2 = Int((Rnd*TotalRecord2))
altreklam.Move(MoveEntry2)
altreklam("show") = altreklam("show") + 1
altreklam.Update
Set reklamalt = Server.CreateObject("ADODB.RecordSet")
bn2SQL = "Select * FROM BANNERS WHERE banner_id="&altreklam("banner_id")&""
reklamalt.open bn2SQL,Connection,3,3
	IF reklamalt("limit") = reklamalt("show") Then
	reklamalt("active") = False
	reklamalt.Update
	ElseIF reklamalt("show") < reklamalt("limit") Then
		IF reklamalt("ban_type") = "Flash" Then
%>
<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" WIDTH="468" HEIGHT="60" onClick="javascript:location.href('redirect.asp?go=ads&bid=<%=reklamalt("banner_id")%>');">
<PARAM NAME="movie" VALUE="<%=reklamalt("ban_url")%>">
<param name="WMode" value="Transparent">
<PARAM NAME="quality" VALUE="high">
<EMBED src="<%=reklamalt("ban_url")%>" quality="high" WIDTH="468" HEIGHT="60"TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED></OBJECT>
<SCRIPT src="inc/mn.js" type=text/javascript></SCRIPT>
<%		ELSEIF reklamalt("ban_type") = "Normal" Then%>
<a href="redirect.asp?go=ads&bid=<%=reklamalt("banner_id")%>" TARGET="_new"><img src="<%=reklamalt("ban_url")%>" width="468" height="60" border="0" align="absmiddle"></a>
<%
		End IF
	End If
reklamalt.close
set reklamalt = nothing
End if
altreklam.close
set altreklam = nothing
%>