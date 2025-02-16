<!--UpdateViewRecord.asp-->
<%@ LANGUAGE = VBScript %>
<!--#INCLUDE FILE="DHFunctions.inc"-->
<HTML>
<HEAD>
<TITLE>Update Viewing Response</TITLE>
</HEAD>
<BODY>
<H1>DreamHome Estate Agents</H1>
<h2>Update Viewing Response</h2>
<%
    Dim comments, wishToRent, SQL, RS
    comments = Request.Form("comment")
    wishToRent = Request.Form("chkRent")
    SQL = "update Viewing set Comment='#1', WishToRent=#2 where ID=" & Request.Form("viewID") & ";"
    SQL = Replace(SQL, "#1", comments)
    If wishToRent = "True" Then
        SQL = Replace(SQL, "#2", "Yes")
    Else
        SQL = Replace(SQL, "#2", "No")
    End If
    ' Response.Write SQL    ' Useful for debugging.
    Application("Conn").Execute(SQL)
    Response.Write "<h3>Comments updated.</h3>"
%>
<a href="ReviewViewings.asp">Return to list of viewings</a>
<a href="index.asp">Home</a>
</BODY>
</HTML>