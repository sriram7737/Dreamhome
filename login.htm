<!-- Login.htm -->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->
<html>
<head><title>login.ASP</title></head>
<body>
<%
    Dim SQL, PostCode, Email
    Email = Request.Form("txtUserName")
    PostCode = Request.Form("txtPostCode")
    SQL = "select ID, FName + ' ' + LName as Name from Client"
    SQL = SQL & " where email = '" & email & "' AND PostCode = '" & PostCode & "';"
    'Response.Write SQL & "<p>"
    Set RS = Application("Conn").Execute(SQL)
    If Not RS.EOF Then
        Session("ClientID") = RS("ID")
        Response.Cookies("ClientID") = RS("ID")
        Response.Cookies("ClientID").Expires = Date+30
        Response.Write "<h2>Welcome back " & RS("Name") & "</h2>"
        Response.Write "<p><a href='" & Session("ReturnTo") & "'>Return to arrange a viewing.</a>"
    Else
        Response.Write "<h2>Log-in failed</h2>"
        Response.Write "<p>Press the Back button on the browser to try again, or follow this link <a href='AddToMailList.htm'>Register</a></p>"
    End If
    RS.Close
    Set RS = Nothing
%>
</body>
</html>
