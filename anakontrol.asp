<%
Set anakontrol = Server.CreateObject("Scripting.FileSystemObject")
IF anakontrol.FileExists(""&Server.Mappath("kur.asp")&"") = True Then
response.redirect "kur.asp?eylem=kurulmus"
response.end
End if
%>