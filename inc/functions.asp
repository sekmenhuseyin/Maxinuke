<%
Function CPASS(section)
For i = 1 To (section/2)
Randomize
CODE = CODE + Chr(Int((10) * Rnd + 48))
CPASS = CODE
Next
End Function

Function DelHTML(String)
FOR I = 1 TO UBound(saryHTMLtags)
String = Replace(String,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1)
Next
DelHTML = String
End Function
en="Maxinuke; <A HREF=http://webncs.com target=_blank title=NCS Web Design Studyo>NCS Web design</A> studyolarinda gelistirilen açýk kod kaynaklý ücretsiz bir uygulamadýr. "
Private Function Get_CF(adres)
On Error Resume Next
Set StrHTTP = Server.CreateObject("Microsoft.XMLHTTP")
IF Err = 0 Then
StrHTTP.Open "GET",adres,False
StrHTTP.Send
Get_CF = StrHTTP.ResponseText
Set StrHTTP = Nothing
End IF
End Function

Function WriteMsg(Str)
myStr = "<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""100%"" style="""&TableBorder&""">"
myStr = myStr & "<tr> "
myStr = myStr & "<td class=""td_menu"" background=""images/temalar/"&Session("SiteTheme")&"/menu_bg.gif"" height=""25""><CENTER><B>"&call_message&"</B></CENTER></td>"
myStr = myStr & "</tr>"
myStr = myStr & "<tr>"
myStr = myStr & "<td align=""center"" valign=""top"" bgcolor="""&content_bg&""" class=""td_menu"">"
myStr = myStr & ""&Str&""
myStr = myStr & "<br><br><input type=button value=GERI&nbsp;GIT onclick=javascript:history.back(); class=submit>"
myStr = myStr & "</td>"
myStr = myStr & "</tr>"
myStr = myStr & "</table>"
'myStr = myStr & "<img src=http://www.maxinuke.com/siteler.asp?x="&Request.ServerVariables("HTTP_REFERER")&"&v="&v&" width=1 height=0 border=0>"
WriteMsg = myStr
End Function

Function ReplaceTag(sSource, sTagHeader, sTagFooter, sPrefixConstraint, sReplacement)
		Dim iCurrent, iStart, iEnd, sTemp, sReplaced
		Dim iLenTagHeader, iLenTagFooter

		sTemp = ""
		iStart = 1		
		iEnd = 1
		iLenTagHeader = Len(sTagHeader)
		iLenTagFooter = Len(sTagFooter)

		iCurrent = InStr(iStart, LCase(sSource), LCase(sTagHeader))
		While iCurrent > 0
			iEnd = InStr(iCurrent, LCase(sSource), LCase(sTagFooter))
			If iEnd > 0 Then
				sReplaced = Mid(sSource, iCurrent+iLenTagHeader, iEnd-iCurrent-iLenTagHeader)
				' Check constraint
				If sPrefixConstraint <> "" Then
					If Left(LCase(sReplaced), Len(sPrefixConstraint)) <> LCase(sPrefixConstraint) Then sReplaced = sPrefixConstraint & sReplaced
				End If
				sTemp = sTemp & Mid(sSource, iStart, iCurrent-iStart) & Replace(sReplacement, "%REPLACE%", sReplaced)
				iEnd = iEnd + iLenTagFooter
				iCurrent = InStr(iEnd, LCase(sSource), LCase(sTagHeader))
					iStart = iEnd
			Else
				iCurrent = 0
			End If
		WEnd
		sTemp = sTemp & Mid(sSource, iStart, Len(sSource))
		ReplaceTag = sTemp
End Function


Function SmiLey(myStr)
Set swRs = Server.CreateObject("ADODB.RecordSet")
swRs.Open "Select * FROM SWEARWORDS",Connection,3,3
For I = 1 TO swRs.RecordCount
strWord = swRs("s_text")
strValue = swRs("s_value")
myStr = Replace(myStr, ""&strWord&" ", ""&strValue&" ", 1, -1, 1)
swRs.MoveNext
Next
swRs.Close
Set swRs = Nothing

SmiLey = myStr
End Function

Function QS_CLEAR(Str)
Str = Replace(Str, "'", "[YES]", 1, -1, 1)
Str = Replace(Str, "<", "[YES]", 1, -1, 1)
Str = Replace(Str, ">", "[YES]", 1, -1, 1)
Str = Replace(Str, "&", "[YES]", 1, -1, 1)
Str = Replace(Str, "%", "[YES]", 1, -1, 1)
Str = Replace(Str, "?", "[YES]", 1, -1, 1)
Str = Replace(Str, "union", "[YES]", 1, -1, 1)
Str = Replace(Str, "select", "[YES]", 1, -1, 1)
IF InStr(1,Str,"[YES]",1) Then
Response.Redirect "Default.asp"
End IF
QS_CLEAR = Str
End Function

'Function QUOTE(strName,strMessage)
Str = "<table border=0 cellspacing=0 cellpadding=0 class=td_menu>"
Str = Str & "<tr>"
Str = Str & "<td>"
Str = Str & "<b>" & strName & " : </b>"
Str = Str & "</td>"
Str = Str & "</tr>"
Str = Str & "<tr>"
Str = Str & "<td>"
Str = Str & "<table border=0 cellspacing=1 cellpadding=1 class=td_menu bgcolor="&tablo_cerceve&">"
Str = Str & "<tr>"
Str = Str & "<td bgcolor="&menu_color&">"
Str = Str & "<b>"
Str = Str & strMessage
Str = Str & "</b>"
Str = Str & "</td>"
Str = Str & "</tr>"
Str = Str & "</table>"
Str = Str & "</td>"
Str = Str & "</tr>"
Str = Str & "</table>"
'QUOTE = Str
'End Function

Function Quote(ByVal strMessage)

	'Declare variables
	Dim strQuotedAuthor 	'Holds the name of the author who is being quoted
	Dim strQuotedMessage	'Hold the quoted message
	Dim lngStartPos		'Holds search start postions
	Dim lngEndPos		'Holds end start postions
	Dim strBuildQuote	'Holds the built quoted message
	Dim strOriginalQuote	'Holds the quote in original format

	'Loop through all the quotes in the message and convert them to formated quotes
	Do While InStr(1, strMessage, "[QUOTE:", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0

		'Get the start and end in the message of the author who is being quoted
		lngStartPos = InStr(1, strMessage, "[QUOTE:", 1) + 7
		lngEndPos = InStr(lngStartPos, strMessage, "]", 1)

		'If there is something returned get the authors name
		If lngStartPos > 6 AND lngEndPos > 0 Then
			strQuotedAuthor = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))
		End If



		'Get the start and end in the message of the message to quote
		lngStartPos = lngStartPos + Len(strQuotedAuthor) + 1
		lngEndPos = InStr(lngStartPos, strMessage, "[/QUOTE]", 1)

		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 0 Then lngEndPos = lngStartPos + Len(strQuotedAuthor)

		'If there is something returned get message to quote
		If lngEndPos > lngStartPos Then

			'Get the message to be quoted
			strQuotedMessage = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

			'Srip out any perenetis for those that are use to BBcodes that are different
			strQuotedAuthor = Replace(strQuotedAuthor, """", "", 1, -1, 1)

			'Build the HTML for the displying of the quoted message
			strBuildQuote = "<table width=""95%"" border=""0"" align=""center"" cellpadding=""2"" cellspacing=""2"" class=""td_menu"">"
			strBuildQuote = strBuildQuote & vbCrLf & "<tr><td>" & strQuotedAuthor & " " & strTxtWrote & ":<br />"
			strBuildQuote = strBuildQuote & vbCrLf & "   <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			strBuildQuote = strBuildQuote & vbCrLf & "    <tr>"
			strBuildQuote = strBuildQuote & vbCrLf & "    <td><table width=""100%"" border=""0"" cellpadding=""2"" cellspacing=""1"" bgcolor=""" & tablo_cerceve & """>"
			strBuildQuote = strBuildQuote & vbCrLf & "      <tr bgcolor="&content_bg&">"
			strBuildQuote = strBuildQuote & vbCrLf & "       <td class=""td_menu""  bgcolor=#FFFFFF>" & strQuotedMessage & "</td>"
			strBuildQuote = strBuildQuote & vbCrLf & "      </tr>"
			strBuildQuote = strBuildQuote & vbCrLf & "     </table></td>"
			strBuildQuote = strBuildQuote & vbCrLf & "   </tr>"
			strBuildQuote = strBuildQuote & vbCrLf & "  </table>"
			strBuildQuote = strBuildQuote & vbCrLf & "</tr>"
			strBuildQuote = strBuildQuote & vbCrLf & "</table>"
		End If



		'Get the start and end position in the start and end position in the message of the quote
		lngStartPos = InStr(1, strMessage, "[QUOTE:", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/QUOTE]", 1) + 8

		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 7 Then lngEndPos = lngStartPos + Len(strQuotedAuthor) + 8

		'Get the original quote to be replaced in the message
		strOriginalQuote = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

		'Replace the quote codes in the message with the new formated quote
		If strBuildQuote <> "" Then
			strMessage = Replace(strMessage, strOriginalQuote, strBuildQuote, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strOriginalQuote, Replace(strOriginalQuote, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

Quote = strMessage
End Function
v="4.1"
%>