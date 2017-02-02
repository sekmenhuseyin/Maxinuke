<!--#include file="Language_File.asp" -->
<!--#include file="database.asp" -->
<%
Session.LCID = 1033
Set sett = Server.CreateObject("ADODB.RecordSet")
settSQL = "Select * FROM SETTINGS"
sett.open settSQL,Connection,1,3 %>
<!--#include file="filter.asp" -->
<!--#include file="themes.asp" -->
<!--#include file="functions.asp" -->
<!--#include file="clear_html_codes.asp" -->