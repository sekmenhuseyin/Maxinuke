<%
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open ("PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" + Server.MapPath("../db/3_5.mdb"))
%>