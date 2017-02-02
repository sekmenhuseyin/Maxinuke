<%
Set sett = Server.CreateObject("ADODB.RecordSet")
settSQL = "Select * FROM SETTINGS"
sett.open settSQL,Connection,1,3

'Session.LCID = sett("s_lcid")
Session.LCID = 1055

IF Session("Enter") = "Yes" Then
Connection.Execute("UPDATE MEMBERS SET u_browser = '"&Browser&"' WHERE uye_id = "&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET u_ws = '"&os&"' WHERE uye_id = "&Session("Uye_ID")&"")
End IF

Set nCount = Server.CreateObject("ADODB.RecordSet")
nSQL = "Select * FROM WEBCOUNTER"
nCount.open nSQL,Connection,1,3
Set c_check = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = date()-0")

If c_check.Eof OR c_check.Bof Then
If Hour(Now()) = 0 OR Hour(Now()) > 0 Then
nCount.AddNew
nCount("today") = "0"
nCount("cdate") = date()
nCount.Update
End If
End If

IF Session("Enter") = "Yes" Then
  Set checkIP = Connection.Execute("SELECT * FROM GUESTS WHERE ip = '"&Request.ServerVariables("REMOTE_ADDR")&"'")
  If Not checkIP.EOF Then
  Set updIP = Connection.Execute("DELETE * FROM GUESTS WHERE ip = '" & Request.ServerVariables("REMOTE_ADDR") & "'")
  End If
  Set updtIP = Connection.Execute("UPDATE MEMBERS SET last_ip = '"&Request.ServerVariables("REMOTE_ADDR")&"' WHERE uye_id = "&Session("Uye_ID")&"")
  Set upddIP = Connection.Execute("UPDATE MEMBERS SET son_tarih = now() WHERE uye_id = "&Session("Uye_ID")&"")
End If

ip = Request.ServerVariables("REMOTE_ADDR")
IF Session("Enter") <> "Yes" Then
Set newguest = Server.CreateObject("ADODB.RecordSet")
guestSQL = "Select * FROM GUESTS"
newguest.open guestSQL,Connection,1,3

dateLimit = DateAdd("n", -1 * 5, Now())
Set checkForIp = Connection.Execute("SELECT * FROM GUESTS WHERE ip = '"& ip &"'")
If Session("Enter") <> "Yes" Then
If checkForIp.EOF Then
Set addIP = Server.CreateObject("ADODB.RecordSet")
addIP.Open "SELECT * FROM GUESTS", Connection, 1, 3
addIP.AddNew
addIP("ip") = ip
addIP("zdate") = Now()
addIP.Update

Else
Set updIP = Server.CreateObject("ADODB.RecordSet")
updIP.Open "SELECT * FROM GUESTS WHERE ip = '"&ip&"'", Connection, 1, 3
updIP("ip") = ip
updIP("zdate") = Now()
updIP.Update
End If
Else
End If
End If

Set delIP = Connection.Execute("DELETE * FROM GUESTS WHERE zdate < now()")

Set LB = Server.CreateObject("ADODB.RecordSet")
LB.open "Select * FROM bloklar WHERE b_align = 'LEFT' ORDER BY b_queue",Connection,3,3

Set RB = Server.CreateObject("ADODB.RecordSet")
RB.open "Select * FROM bloklar WHERE b_align = 'RIGHT' ORDER BY b_queue",Connection,3,3

IF Session("Enter") = "Yes" Then
Session("SiteTheme") = Session("Theme")
ELSE
Session("SiteTheme") = sett("theme")
End IF
%>