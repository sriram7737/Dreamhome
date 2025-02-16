<!--createBooking.asp-->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->
<html>
<head><title>Create a new booking</title></head>

<body>
<%
    Dim SQL
    If Session("propID") <> "" Then
        SQL = "insert into Viewing(clientID, propertyNo, viewDate, viewHour) values(#c, '#p', '#d', #h);"
        SQL = Replace(SQL, "#c", Session("ClientID"))
        SQL = Replace(SQL, "#p", Session("propID"))
        SQL = Replace(SQL, "#d", Request.Form("optDate"))
        SQL = Replace(SQL, "#h", Request.Form("optHour"))
        Application("Conn").Execute SQL
        Response.Write "<h2>Booking created</h2>"
        Response.Write "<h3>Property: " & Session("propID") & "</h3>"
        Response.Write "<h3>Date: " & Request.Form("optDate") & "</h3>"
        Response.Write "<h3>Hour: " & Request.Form("optHour") & "</h3>"
        Response.Write "<a href='index.asp'>Home</a>"
    Else
        Response.Write "No property has been selected."
        Response.Write "<a href='index.asp'>Home</a>"
    End If
%>
</body>
</html>