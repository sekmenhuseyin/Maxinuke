<%
veritabani_adi="maxinuke"
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/"&veritabani_adi&".mdb")
%>