<!-- This sequence will start every page. -->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->
<html>
<head>
<title>DreamHome Estate Agents</title>
</head>
<body>
<h1>DreamHome Estate Agents</h1>
<%
	' Do we know this client...
	If Session("ClientID") <> "" Then
		Name = GetClientName(Session("ClientID"))
		Response.Write "<h3>Hello " & Name & "</h3>"
		Response.Write "<p>Welcome back to the Dreamhome website.</p>"	
        Response.Write "<p><a href='rentals.htm'>Properties for Rent.</a></p>"
        Response.Write "<p><a href='ReviewViewings.asp'>" & _
                            "Check your viewing appointments.</a></p>"
	Else
             Response.Write "<p><a href='rentals.htm'>Properties for Rent.</a></p>"
		     Response.Write "<p><a href='AddToMailList.htm'>" & _
                            "Join our Mailing List</a></p>"
	End If
%>
</body>
</html>
</html>