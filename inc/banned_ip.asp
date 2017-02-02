<%
Set bipRs = Server.CreateObject("ADODB.RecordSet")
bipRs.open "Select * FROM BANNED_IPS",Connection,3,3
Do Until bipRs.EoF
IF bipRs("b_ip") = Request.ServerVariables("REMOTE_ADDR") OR bipRs("b_ip") = ""&Session("Uye_ID")&"" Then
Response.Write "<div align=""center"" class=""td_menu"" style=""font-size:14px""><b>" & banned_ip_text & "</b></div>"
Response.End
End IF
bipRs.MoveNext
Loop
bipRs.Close
Set bipRs = Nothing
%>