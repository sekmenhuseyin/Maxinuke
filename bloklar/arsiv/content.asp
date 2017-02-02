<!--#include file="../../inc/bloklar_includes.asp" -->
<!--
#####################################################################################################################
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#												Arsiv Blok Uygulamasi												#
#----------------------------------------------------- MAXI NUKE ---------------------------------------------------#
#####################################################################################################################
-->
<%
day1 = "Pzt"
day2 = "Sl"
day3 = "Çþ"
day4 = "Pþ"
day5 = "Cm"
day6 = "Ct"
day7 = "Pzr"
m1 = "Ocak"
m2 = "Þubat"
m3 = "Mart"
m4 = "Nisan"
m5 = "Mayýs"
m6 = "Haziran"
m7 = "Temmuz"
m8 = "Aðustos"
m9 = "Eylül"
m10 = "Ekim"
m11 = "Kasým"
m12 = "Aralýk"
ay	=	Request.QueryString("month")
yil	=	Year(Now())
if ay = "" then ay = month(date)
if yil= "" then yil= year(date)
	if ay > 12 Then 
	ay = 1
	yil = yil+1
	end if
	if ay < 1 then
	ay = 12
	yil = yil-1
	end if
gun=28
while isdate(gun&","&ay&","&yil)=true
gun = gun+1
wend
gun = gun-1
If ay = 1 Then
ayy = m1
elseIf ay = 2 Then
ayy = m2
elseIf ay = 3 Then
ayy = m3
elseIf ay = 4 Then
ayy = m4
elseIf ay = 5 Then
ayy = m5
elseIf ay = 6 Then
ayy = m6
elseIf ay = 7 Then
ayy = m7
elseIf ay = 8 Then
ayy = m8
elseIf ay = 9 Then
ayy = m9
elseIf ay = 10 Then
ayy = m10
elseIf ay = 11 Then
ayy = m11
elseIf ay = 12 Then
ayy = m12
End If
%>
<form method="GET" action="#" name="form1">
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="td_menu" align="center">
	<tr>
		<td><a href="calender.asp?action=show&day=1&month=<%=ay-1%>&year=<%=yil%>">«</a></td>
		<td align="center" colspan="5"><b><%=ayy%></b></td>
		<td><a href="calender.asp?action=show&day=1&month=<%=ay+1%>&year=<%=yil%>">»</a></td>
	</tr>
	<tr>
		<td><%=day1%></td>
		<td><%=day2%></td>
		<td><%=day3%></td>
		<td><%=day4%></td>
		<td><%=day5%></td>
		<td><%=day6%></td>
		<td><%=day7%></td>
	</tr>
	<tr>
		<td colspan="7"><hr size="1" color="<%=tablo_cerceve%>"></td>
	</tr>
	<tr>
<%
sayac=weekday("01,"&ay&","&yil)
if sayac=1 then
sayac=6
else
sayac=sayac-5
end if
for i=1 to sayac
Response.Write("<td> </td>")
next
for i=1 to gun
if i = Day(Now()) then
a = i
else
a = i
end if
if sayac mod 7=0 then Response.Write("</tr> <tr>")%>
<td <% IF a = Day(Now()) Then%>style="<%=TableBorder%>"<% End IF %>><a href="calender.asp?action=show&day=<%=a%>&month=<%=ay%>"><% IF a = Day(Now()) Then%><b><%=a%></b><%Else%><%=a%><% End IF %></a></td>
<%
sayac=sayac+1
next
%>
</tr>
</table>
<div align="center">
<select name="month" class="form1">
	<option value="01"><%=m1%></option>
	<option value="02"><%=m2%></option>
	<option value="03"><%=m3%></option>
	<option value="04"><%=m4%></option>
	<option value="05"><%=m5%></option>
	<option value="06"><%=m6%></option>
	<option value="07"><%=m7%></option>
	<option value="08"><%=m8%></option>
	<option value="09"><%=m9%></option>
	<option value="10"><%=m10%></option>
	<option value="11"><%=m11%></option>
	<option value="12"><%=m12%></option>
</select>	
<input type="submit" name="button" value="<%=calender_go%>" class="submit">
</div>
</form>