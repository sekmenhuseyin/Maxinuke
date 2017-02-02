<!--#include file="ust.asp" -->
<!--#include file="view.asp" -->
<%
call ORTA
page_id = Request.QueryString("id")
page_id = QS_CLEAR(page_id)
Set pageRs = Server.CreateObject("ADODB.RecordSet")
pageRs.open "SELECT * FROM PAGES WHERE p_id = "&page_id&"",Connection,3,3
IF pageRs.Eof OR PageRs.Bof Then
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""0"" align=""center"" width=""100%"" class=""td_menu"" bgcolor="""&tablo_cerceve&""" height=""20""><tr><td align=""center"" bgcolor="""&content_bg&"""><b>"&module_eof&"</b></td></tr></table>"
ELSE
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="images/temalar/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><CENTER><B><%=pageRs("p_title")%></B></CENTER></td>
                </tr>
                <tr> 
<td align="left" valign="top" bgcolor="<%=content_bg%>" class="td_menu">
<%
page_content = pageRs("p_content")
page_content = Replace(page_content, vbCrLf, "<br>", 1, -1, 1)
If pageRs("mems") = True Then
If Session("Enter") = "Yes" Then
Response.Write page_content
Else
Response.Write "<center>"&sett("np_msg")&"</center>"
End If
ELSE
Response.Write page_content
End IF
%>
</td>
</tr>
</table>
<%
End IF
call ORTA2
call ALT %>