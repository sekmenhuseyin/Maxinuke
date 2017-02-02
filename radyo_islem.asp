<%
If Request.QueryString("eylem")="ac" Then
session("radyo")="gorun"
Response.Redirect "default.asp"
elseIf Request.QueryString("eylem")="kapa" Then
session("radyo")="gizlen"
Response.Redirect "default.asp"
End If
%>