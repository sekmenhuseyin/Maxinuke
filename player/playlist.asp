<!--#include file="bag.asp"-->
<Asx Version="3.0">
<Param Name="AllowShuffle" Value="yes"/>
<%
Set muzik = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * from muzikler"
muzik.open SQL,Connection,1,3
for k=1 to muzik.pagesize
if muzik.eof or muzik.bof then exit for
%>
<Entry>
	<title><%=muzik("sanatci")%> - <%=muzik("parca")%></title>
<%
If muzik("tip") = "1" Then
%>
	<Ref href="http://127.0.0.1/269/uploads/mp3/<%=muzik("dosya_adi")%>"/>
<%
elseIf muzik("tip") = "2" Then
%>
	<Ref href="<%=muzik("dosya_adi")%>"/>
<%End if%>
</Entry>
<%
muzik.movenext : next
muzik.Close
Set muzik = Nothing
%>
</Asx>