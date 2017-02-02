<%
Set ustreklam = Server.CreateObject("ADODB.Recordset" )
b1SQL = "Select * FROM BANNERS WHERE active=True AND position='top'"
ustreklam.Open b1SQL,Connection,3,3
if ustreklam.eof then
%>
<!--#include file="inc/adsense_3.asp" -->
<%
else
TotalRecord1 = ustreklam.RecordCount
Randomize
MoveEntry1 = Int((Rnd*TotalRecord1))
ustreklam.Move(MoveEntry1)
ustreklam("show") = ustreklam("show") + 1
ustreklam.Update
Set reklamust = Server.CreateObject("ADODB.RecordSet")
bn1SQL = "Select * FROM BANNERS WHERE banner_id="&ustreklam("banner_id")&""
reklamust.open bn1SQL,Connection,3,3
	IF reklamust("limit") = reklamust("show") Then
	reklamust("active") = False
	reklamust.Update
	ElseIF reklamust("show") < reklamust("limit") Then
		IF reklamust("ban_type") = "Flash" Then
%>
<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" WIDTH="468" HEIGHT="60" onClick="javascript:location.href('redirect.asp?go=ads&bid=<%=reklamust("banner_id")%>');">
<PARAM NAME="movie" VALUE="<%=reklamust("ban_url")%>">
<param name="WMode" value="Transparent">
<PARAM NAME="quality" VALUE="high">
<EMBED src="<%=reklamust("ban_url")%>" quality="high" WIDTH="468" HEIGHT="60"TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED></OBJECT>
<SCRIPT src="inc/mn.js" type=text/javascript></SCRIPT>
<%		ELSEIF reklamust("ban_type") = "Normal" Then%>
<a href="redirect.asp?go=ads&bid=<%=reklamust("banner_id")%>" TARGET="_new"><img src="<%=reklamust("ban_url")%>" width="468" height="60" border="0" align="absmiddle"></a>
<%
		End IF
	End If
reklamust.close
set reklamust = nothing
End if
ustreklam.close
set ustreklam = nothing
%>