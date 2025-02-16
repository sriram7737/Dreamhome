<!--ReviewViewings.asp-->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->
<HTML>
<HEAD>
<TITLE>Review Property Viewings</TITLE>
</HEAD>
<BODY>
<H1>DreamHome Estate Agents</H1>
<h2>Review Property Viewings</h2><br>
<%
    Response.Write "<h3>Client : " & GetClientName(Session("ClientID")) & "</h3>"
    Dim RS 
    Dim SQL
    ' Set up a SQL statement to retrieve the current cllient's viewing records...
    SQL = "select ID, propertyNo, viewDate, viewHour, Comment, WishToRent from Viewing where clientID=" & Session("ClientID")
    Set RS = Application("Conn").Execute(SQL)
    ' Write out the result as a HTML table...
    Response.Write RSToHTMLTable(RS, 1, "ViewingComments.asp")
    If RS.EOF Then
        Response.Write "<br>You have not viewed any properties.<br>"
    End If
%>
<a href="index.asp">Home</a>
</BODY>
</HTML>
