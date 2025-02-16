<!--bookViewing.asp-->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->
<html>
<head><title>Book A Viewing</title></head>

<body>
<%
    ' First we need to establish who the booking is for...
    If Session("ClientID") = "" Then
        ' Don't know who - the client will need to log-in.
        ' Start by storing the url used to access this page so we can easily return to it...
        Session("ReturnTo") = Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
        'Response.Write Session("ReturnTo")
        Response.Redirect "login.htm"
        'Response.Write "<br><a href='login.htm'>Log-in to site</a>"
    Else
        ' We know this client, so can continue with the booking...
        Dim propID, SQL, RS
        Session("ReturnTo") = ""
        propID = Request.QueryString("propID")
        Session("propID") = propID
        SQL = "select * from PropertyForRent inner join Staff " & _
              "on PropertyForRent.StaffNo=Staff.StaffNo " & _
              "where PropertyForRent.PropertyNo='" & propID & "';"
        'Response.Write SQL
        Set RS = Application("Conn").Execute(SQL)
        Response.Write "<h2>Book a viewing</h2>"
        Response.Write "<form method='POST' action='createBooking.asp'>"
        Response.Write "<h3>Property: " & RS("street") & ", " & RS("city") & ", " & RS("PostCode") & "</h3>"
        ' Response.Write "Book a viewing for property number " & propID & "<br>"
        Response.Write "<h3>Client Name: " & GetClientName(Session("ClientID")) & "</h3>"
        Response.Write "<h3>Agent Name: " & GetStaffName(RS("Staff.StaffNo")) & "</h3>"
        Response.Write "<h3>Viewing Date: " & GenerateDateOptions(Date+1, Date+8) & "</h3>"
        Response.Write "<h3>Preferred Hour: " & GenerateHourOptions(9, 5) & "</h3>"
        Response.Write "<br><input type='submit' value='Submit'><input type='button' value='Cancel'></form>"
    End If
%>
</body>
</html>
